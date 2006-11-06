VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MWCapillaryFlowClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Molecular Weight Calculator routines with ActiveX Class interfaces: MWCapillaryFlowClass
' Written by Matthew Monroe in Richland, WA

' <EnumStatements>
Public Enum ctCapillaryTypeConstants
    ctOpenTubularCapillary = 0
    ctPackedCapillary
End Enum

Public Enum uprUnitsPressureConstants
    uprPsi = 0
    uprPascals
    uprKiloPascals
    uprAtmospheres
    uprBar
    uprTorr
    uprDynesPerSquareCm
End Enum

Public Enum ulnUnitsLengthConstants
    ulnM = 0
    ulnCM
    ulnMM
    ulnMicrons
    ulnInches
End Enum

Public Enum uviUnitsViscosityConstants
    uviPoise = 0
    uviCentiPoise
End Enum

Public Enum ufrUnitsFlowRateConstants
    ufrMLPerMin = 0
    ufrULPerMin
    ufrNLPerMin
End Enum

Public Enum ulvUnitsLinearVelocityConstants
    ulvCmPerHr = 0
    ulvMmPerHr
    ulvCmPerMin
    ulvMmPerMin
    ulvCmPerSec
    ulvMmPerSec
End Enum

Public Enum utmUnitsTimeConstants
    utmHours = 0
    utmMinutes
    utmSeconds
End Enum

Public Enum uvoUnitsVolumeConstants
    uvoML = 0
    uvoUL
    uvoNL
    uvoPL
End Enum

Public Enum ucoUnitsConcentrationConstants
    ucoMolar = 0
    ucoMilliMolar
    ucoMicroMolar
    ucoNanoMolar
    ucoPicoMolar
    ucoFemtoMolar
    ucoAttoMolar
    ucoMgPerML
    ucoUgPerML
    ucongperml
    ucoUgPerUL
    ucoNgPerUL
End Enum

Public Enum utpUnitsTemperatureConstants
    utpCelsius = 0
    utpKelvin
    utpFahrenheit
End Enum

Public Enum umfMassFlowRateConstants
    umfPmolPerMin = 0
    umfFmolPerMin
    umfAmolPerMin
    umfPmolPerSec
    umfFmolPerSec
    umfAmolPerSec
    umfMolesPerMin
End Enum

Public Enum umaMolarAmountConstants
    umaMoles = 0
    umaMilliMoles
    umaMicroMoles
    umaNanoMoles
    umaPicoMoles
    umaFemtoMoles
    umaAttoMoles
End Enum

Public Enum udcDiffusionCoefficientConstants
    udcCmSquaredPerHr = 0
    udcCmSquaredPerMin
    udcCmSquaredPerSec
End Enum

Public Enum acmAutoComputeModeConstants
    acmBackPressure = 0
    acmColumnID
    acmColumnLength
    acmDeadTime
    acmLinearVelociy
    acmVolFlowrate
    acmVolFlowrateUsingDeadTime
End Enum
' </EnumStatements>

' <UDT's>
Private Type udtCapillaryFlowParametersType
    CapillaryType As ctCapillaryTypeConstants
    BackPressure As Double              ' in dynes/cm^2
    ColumnLength As Double              ' in cm
    ColumnID As Double                  ' in cm
    SolventViscosity As Double          ' in poise
    ParticleDiameter As Double          ' in cm
    VolumetricFlowRate As Double        ' in mL/min
    LinearVelocity As Double            ' cm/min
    ColumnDeadTime As Double            ' in min
    InterparticlePorosity As Double
End Type

Private Type udtMassRateParametersType
    SampleConcentration As Double       ' in Molar
    SampleMass As Double                ' in g/mole
    VolumetricFlowRate As Double        ' in mL/min
    InjectionTime As Double             ' in min
    MassFlowRate As Double              ' in Moles/min
    MolesInjected As Double             ' in moles
End Type

Private Type udtExtraColumnBroadeningParametersType
    LinearVelocity As Double                ' in cm/min
    DiffusionCoefficient As Double          ' in cm^2/sec
    OpenTubeLength As Double                ' in cm
    OpenTubeID As Double                    ' in cm
    InitialPeakWidth As Double              ' in sec
    TemporalVariance As Double              ' in sec^2
    AdditionalTemporalVariance As Double    ' in sec^2
    ResultantPeakWidth As Double            ' in sec
End Type
' </UDT's>

' Conversion Factors
Private Const CM_PER_INCH As Single = 2.54
Private Const PI As Double = 3.14159265359

Private mCapillaryFlowParameters As udtCapillaryFlowParametersType
Private mMassRateParameters As udtMassRateParametersType
Private mExtraColumnBroadeningParameters As udtExtraColumnBroadeningParametersType

Private mAutoCompute As Boolean                           ' When true, automatically compute results whenever any value changes
Private mAutoComputeMode As acmAutoComputeModeConstants   ' The value to compute when mAutoCompute is true

Private Sub CheckAutoCompute()
    If mAutoCompute Then
        Select Case mAutoComputeMode
        Case acmBackPressure: ComputeBackPressure
        Case acmColumnID: ComputeColumnID
        Case acmColumnLength: ComputeColumnLength
        Case acmDeadTime: ComputeDeadTime
        Case acmLinearVelociy: ComputeLinearVelocity
        Case acmVolFlowrateUsingDeadTime: ComputeVolFlowRateUsingDeadTime
        Case Else
            ' Includes acmVolFlowRate
            ComputeVolFlowRate
        End Select
    End If
End Sub

Public Function ComputeBackPressure(Optional eUnits As uprUnitsPressureConstants = uprPsi, Optional blnComputeDeadTime As Boolean = True) As Double
    ' Computes the back pressure, stores in .BackPressure, and returns it
    
    Dim dblBackPressure As Double, dblRadius As Double
    
    With mCapillaryFlowParameters
        
        dblRadius = .ColumnID / 2#
        
        If dblRadius <> 0 Then
            If .CapillaryType = ctOpenTubularCapillary Then
                ' Open tubular capillary
                dblBackPressure = (.VolumetricFlowRate * 8 * .SolventViscosity * .ColumnLength) / (dblRadius ^ 4 * PI * 60) ' Pressure in dynes/cm^2
            Else
                ' Packed capillary
                If .ParticleDiameter <> 0 And .InterparticlePorosity <> 0 Then
                    ' Flow rate in mL/sec
                    dblBackPressure = (.VolumetricFlowRate * 180 * .SolventViscosity * .ColumnLength * (1 - .InterparticlePorosity) ^ 2) / _
                                      (.ParticleDiameter ^ 2 * .InterparticlePorosity ^ 2 * PI * dblRadius ^ 2 * 60) / .InterparticlePorosity
                Else
                    dblBackPressure = 0
                End If
            End If
        Else
            dblBackPressure = 0
        End If
        
        .BackPressure = dblBackPressure
    End With
    
    If blnComputeDeadTime Then
        ComputeDeadTime utmMinutes, False
    End If
    
    ' Compute Linear Velocity
    ComputeLinearVelocity ulvCmPerSec, False
    
    ' Return Back Pressure
    ComputeBackPressure = ConvertPressure(dblBackPressure, uprDynesPerSquareCm, eUnits)
    
End Function

Public Function ComputeColumnLength(Optional eUnits As ulnUnitsLengthConstants = ulnCM) As Double
    ' Computes the column length, stores in .ColumnLength, and returns it
    
    Dim dblColumnLength As Double, dblRadius As Double
    
    With mCapillaryFlowParameters
        
        dblRadius = .ColumnID / 2#
        
        If .SolventViscosity <> 0 And .VolumetricFlowRate <> 0 Then
            If .CapillaryType = ctOpenTubularCapillary Then
                ' Open tubular capillary
                dblColumnLength = (.BackPressure * dblRadius ^ 4 * PI * 60) / (8 * .SolventViscosity * .VolumetricFlowRate) ' Column length in cm
            Else
                ' Packed capillary
                If .InterparticlePorosity <> 1 Then
                    ' Flow rate in mL/sec
                    dblColumnLength = (.BackPressure * .ParticleDiameter ^ 2 * .InterparticlePorosity ^ 2 * PI * dblRadius ^ 2 * 60) * .InterparticlePorosity / _
                              (180 * .SolventViscosity * .VolumetricFlowRate * (1 - .InterparticlePorosity) ^ 2)
                Else
                    dblColumnLength = 0
                End If
            End If
        Else
            dblColumnLength = 0
        End If
        
        .ColumnLength = dblColumnLength
    End With
    
    ' Compute Dead Time (and Linear Velocity)
    ComputeDeadTime utmMinutes, False
    
    ' Return Column Length
    ComputeColumnLength = ConvertLength(dblColumnLength, ulnCM, eUnits)
    
End Function

Public Function ComputeColumnVolume(Optional eUnits As uvoUnitsVolumeConstants) As Double
    ' Computes the column volume and returns it (does not store it)
    
    Dim dblColumnVolume As Double, dblRadius As Double
    
    With mCapillaryFlowParameters
        
        dblRadius = .ColumnID / 2#
    
        dblColumnVolume = .ColumnLength * PI * dblRadius ^ 2    ' In mL
        
        If .CapillaryType = ctPackedCapillary Then
            dblColumnVolume = dblColumnVolume * .InterparticlePorosity
        End If
    End With
    
    ComputeColumnVolume = ConvertVolume(dblColumnVolume, uvoML, eUnits)
    
End Function

Public Function ComputeColumnID(Optional eUnits As ulnUnitsLengthConstants = ulnMicrons) As Double
    ' Computes the column length, stores in .ColumnLength, and returns it
    
    Dim dblRadius As Double
    
    With mCapillaryFlowParameters
        
        If .BackPressure <> 0 Then
            If .CapillaryType = ctOpenTubularCapillary Then
                ' Open tubular capillary
                dblRadius = ((.VolumetricFlowRate * 8 * .SolventViscosity * .ColumnLength) / _
                            (.BackPressure * PI * 60)) ^ (0.25)
            Else
                ' Packed capillary
                If .ParticleDiameter <> 0 And .InterparticlePorosity <> 1 And .InterparticlePorosity <> 0 Then
                    ' Flow rate in mL/sec
                    dblRadius = ((.VolumetricFlowRate * 180 * .SolventViscosity * .ColumnLength * (1 - .InterparticlePorosity) ^ 2) / _
                                (.BackPressure * .ParticleDiameter ^ 2 * .InterparticlePorosity ^ 2 * PI * 60) / .InterparticlePorosity) ^ 0.5
                Else
                    dblRadius = 0
                End If
            End If
        Else
            dblRadius = 0
        End If
        
        .ColumnID = dblRadius * 2#
    End With
    
    ' Compute Dead Time (and Linear Velocity)
    ComputeDeadTime utmMinutes, False
    
    ' Return Column ID
    ComputeColumnID = ConvertLength(dblRadius * 2#, ulnCM, eUnits)
    
End Function

Public Function ComputeDeadTime(Optional eUnits As utmUnitsTimeConstants = utmMinutes, Optional blnRecalculateVolFlowRate As Boolean = True) As Double
    ' Computes the column dead time, stores in .ColumnDeadTime, and returns it
    
    Dim dblDeadTime As Double
    
    ' Dead time is dependent on Linear Velocity, so compute
    ComputeLinearVelocity ulvCmPerSec, blnRecalculateVolFlowRate
    
    With mCapillaryFlowParameters
        
        If .LinearVelocity <> 0 Then
            dblDeadTime = .ColumnLength / .LinearVelocity   ' Dead time in minutes
        Else
            dblDeadTime = 0
        End If
        
        .ColumnDeadTime = dblDeadTime
    End With
    
    ' Return Dead Time
    ComputeDeadTime = ConvertTime(mCapillaryFlowParameters.ColumnDeadTime, utmMinutes, eUnits)
    
End Function

Public Function ComputeExtraColumnBroadeningResultantPeakWidth(Optional eUnits As utmUnitsTimeConstants = utmSeconds) As Double
    ComputeExtraColumnBroadeningValues
    
    ComputeExtraColumnBroadeningResultantPeakWidth = GetExtraColumnBroadeningResultantPeakWidth(eUnits)
End Function

Private Sub ComputeExtraColumnBroadeningValues()
    Dim dblInitialPeakVariance As Double
    Dim dblSumOfVariances As Double
    
    With mExtraColumnBroadeningParameters
        If .LinearVelocity <> 0 And .DiffusionCoefficient <> 0 Then
            .TemporalVariance = .OpenTubeID ^ 2 * .OpenTubeLength / (96 * .DiffusionCoefficient * .LinearVelocity / 60) ' in sec^2
        Else
            .TemporalVariance = 0
        End If
        
        dblInitialPeakVariance = (.InitialPeakWidth / 4) ^ 2
        
        dblSumOfVariances = dblInitialPeakVariance + .TemporalVariance + .AdditionalTemporalVariance
        
        If dblSumOfVariances >= 0 Then
            ' ResultantPeakWidth at the base = 4 sigma  and  sigma = Sqr(Total_Variance)
            .ResultantPeakWidth = 4 * Sqr(dblSumOfVariances)
        Else
            .ResultantPeakWidth = 0
        End If
    End With
End Sub

Public Function ComputeLinearVelocity(Optional eUnits As ulvUnitsLinearVelocityConstants = ulvCmPerSec, Optional blnRecalculateVolFlowRate As Boolean = True) As Double
    ' Computes the Linear velocity, stores in .LinearVelocity, and returns it
    
    Dim dblLinearVelocity As Double, dblRadius As Double
    
    If blnRecalculateVolFlowRate Then
        ComputeVolFlowRate ufrMLPerMin
    End If
    
    With mCapillaryFlowParameters
        dblRadius = .ColumnID / 2#
        If dblRadius <> 0 Then
            dblLinearVelocity = .VolumetricFlowRate / (PI * dblRadius ^ 2)    ' Units in cm/min
            
            ' Divide Linear Velocity by epsilon if a packed capillary
            If .CapillaryType = ctPackedCapillary And .InterparticlePorosity <> 0 Then
                dblLinearVelocity = dblLinearVelocity / .InterparticlePorosity
            End If
        Else
            dblLinearVelocity = 0
        End If
        
        .LinearVelocity = dblLinearVelocity
    End With
        
    ' Return Linear Velocity
    ComputeLinearVelocity = ConvertLinearVelocity(dblLinearVelocity, ulvCmPerMin, eUnits)

End Function

Public Function ComputeMassFlowRate(Optional eUnits As umfMassFlowRateConstants = umfFmolPerSec) As Double
    ' Computes the MassFlowRate and Moles Injected, stores in .MassFlowRate and .MolesInjected, and returns MassFlowRate
    
    ComputeMassRateValues
    ComputeMassFlowRate = GetMassFlowRate(eUnits)
    
End Function

Public Function ComputeMassRateMolesInjected(Optional eUnits As umaMolarAmountConstants = umaFemtoMoles) As Double
    ' Computes the MassFlowRate and Moles Injected, stores in .MassFlowRate and .MolesInjected, and returns MassFlowRate
    
    ComputeMassRateValues
    ComputeMassRateMolesInjected = GetMassRateMolesInjected(eUnits)
    
End Function

Private Sub ComputeMassRateValues()
    
    With mMassRateParameters
        .MassFlowRate = .SampleConcentration * .VolumetricFlowRate / 1000   ' Compute mass flowrate in moles/min
        
        .MolesInjected = .MassFlowRate * .InjectionTime                     ' Compute moles injected in moles
    End With
    
End Sub

Public Function ComputeOptimumLinearVelocityUsingParticleDiamAndDiffusionCoeff(Optional eUnits As ulvUnitsLinearVelocityConstants = ulvCmPerSec) As Double
    ' Computes the optimum linear velocity, based on
    ' mCapillaryFlowParameters.ParticleDiameter
    ' and mExtraColumnBroadeningParameters.DiffusionCoefficient
    
    Dim dblOptimumLinearVelocity As Double
    
    With mCapillaryFlowParameters
        If .ParticleDiameter <> 0 Then
            dblOptimumLinearVelocity = 3 * mExtraColumnBroadeningParameters.DiffusionCoefficient / .ParticleDiameter
        
            dblOptimumLinearVelocity = ConvertLinearVelocity(dblOptimumLinearVelocity, ulvCmPerSec, eUnits)
        End If
    End With
    
    ComputeOptimumLinearVelocityUsingParticleDiamAndDiffusionCoeff = dblOptimumLinearVelocity

End Function

Public Function ComputeMeCNViscosity(dblPercentAcetonitrile As Double, dblTemperature As Double, Optional eTemperatureUnits As utpUnitsTemperatureConstants = utpCelsius, Optional eViscosityUnits As uviUnitsViscosityConstants = uviPoise) As Double
    
    Dim dblPhi As Double            ' Fraction Acetonitrile
    Dim dblKelvin As Double
    Dim dblViscosityInCentiPoise As Double
    
    On Error GoTo ComputeMeCNViscosityErrorHandler
    
    dblPhi = dblPercentAcetonitrile / 100#
    If dblPhi < 0 Then dblPhi = 0
    If dblPhi > 100 Then dblPhi = 100
    
    dblKelvin = ConvertTemperature(dblTemperature, eTemperatureUnits, utpKelvin)
    
    If dblKelvin > 0 Then
        dblViscosityInCentiPoise = Exp(dblPhi * (-3.476 + 726 / dblKelvin) + (1 - dblPhi) * (-5.414 + 1566 / dblKelvin) + dblPhi * (1 - dblPhi) * (-1.762 + 929 / dblKelvin))
    Else
        dblViscosityInCentiPoise = 0
    End If
    
    ComputeMeCNViscosity = ConvertViscosity(dblViscosityInCentiPoise, uviCentiPoise, eViscosityUnits)
    Exit Function

ComputeMeCNViscosityErrorHandler:
    ComputeMeCNViscosity = 0
    
End Function


Public Function ComputeVolFlowRate(Optional eUnits As ufrUnitsFlowRateConstants = ufrNLPerMin) As Double
    ' Computes the Volumetric flow rate, stores in .VolumetricFlowRate, and returns it
    
    Dim dblVolFlowRate As Double, dblRadius As Double
    
    With mCapillaryFlowParameters
        
        dblRadius = .ColumnID / 2#
        
        If .SolventViscosity <> 0 And .ColumnLength <> 0 Then
            If .CapillaryType = ctOpenTubularCapillary Then
                ' Open tubular capillary
                dblVolFlowRate = (.BackPressure * dblRadius ^ 4 * PI) / (8 * .SolventViscosity * .ColumnLength)    ' Flow rate in mL/sec
            Else
                ' Packed capillary
                If .InterparticlePorosity <> 1 Then
                    ' Flow rate in mL/sec
                    dblVolFlowRate = (.BackPressure * .ParticleDiameter ^ 2 * .InterparticlePorosity ^ 2 * PI * dblRadius ^ 2) * .InterparticlePorosity / _
                                     (180 * .SolventViscosity * .ColumnLength * (1 - .InterparticlePorosity) ^ 2)
                Else
                    dblVolFlowRate = 0
                End If
            End If
            
            ' Convert dblVolFlowRate to mL/min
            dblVolFlowRate = dblVolFlowRate * 60
        Else
            dblVolFlowRate = 0
        End If
        
        .VolumetricFlowRate = dblVolFlowRate
    End With
    
    ' Compute Dead Time (and Linear Velocity)
    ComputeDeadTime utmMinutes, False
    
    ComputeVolFlowRate = ConvertVolFlowRate(dblVolFlowRate, ufrMLPerMin, eUnits)
End Function

Public Function ComputeVolFlowRateUsingDeadTime(Optional eUnits As ufrUnitsFlowRateConstants = ufrNLPerMin, Optional ByRef dblNewBackPressure As Double, Optional ePressureUnits As uprUnitsPressureConstants = uprPsi) As Double
    ' Computes the Volumetric flow rate using the dead time, stores in .VolumetricFlowRate, and returns it
    ' This requires modifying the pressure value to give the computed volumetric flow rate
    
    Dim dblVolFlowRate As Double, dblRadius As Double
    
    With mCapillaryFlowParameters
        
        dblRadius = .ColumnID / 2#
        
        ' First find vol flow rate that gives observed dead time
        If .ColumnDeadTime <> 0 Then
            
            dblVolFlowRate = .ColumnLength * (PI * dblRadius ^ 2) / .ColumnDeadTime   ' Vol flow rate in mL/sec
            
            If .CapillaryType = ctPackedCapillary Then
                ' Packed Capillary
                dblVolFlowRate = dblVolFlowRate * .InterparticlePorosity
            End If
            
            ' Store the new value
            .VolumetricFlowRate = dblVolFlowRate
            
            ' Now find pressure that gives computed dblVolFlowRate
            ' The ComputeBackPressure sub will store the new pressure
            dblNewBackPressure = ComputeBackPressure(ePressureUnits, False)
        Else
            dblVolFlowRate = 0
            .VolumetricFlowRate = 0
        End If
        
    End With
    
    ' Compute Linear Velocity (but not the dead time)
    ComputeLinearVelocity ulvCmPerSec, False
    
    ComputeVolFlowRateUsingDeadTime = ConvertVolFlowRate(dblVolFlowRate, ufrMLPerMin, eUnits)
End Function

' Duplicated function, in both MWCapillaryFlowClass and MWMoleMassDilutionClass
Public Function ConvertConcentration(dblConcentrationIn As Double, eCurrentUnits As ucoUnitsConcentrationConstants, eNewUnits As ucoUnitsConcentrationConstants) As Double
    Dim dblValue As Double, dblFactor As Double
    Dim dblSampleMass As Double
    
    If eCurrentUnits = eNewUnits Then
        ConvertConcentration = dblConcentrationIn
        Exit Function
    End If
    
    dblSampleMass = mMassRateParameters.SampleMass
    
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

Public Function ConvertDiffusionCoefficient(dblDiffusionCoefficientIn As Double, eCurrentUnits As udcDiffusionCoefficientConstants, eNewUnits As udcDiffusionCoefficientConstants) As Double
    Dim dblValue As Double, dblFactor As Double
    
    If eCurrentUnits = eNewUnits Then
        ConvertDiffusionCoefficient = dblDiffusionCoefficientIn
        Exit Function
    End If
    
    dblFactor = FactorDiffusionCoeff(eCurrentUnits)
    If dblFactor = -1 Then
        ConvertDiffusionCoefficient = -1
        Exit Function
    Else
        dblValue = dblDiffusionCoefficientIn * dblFactor
    End If
    
    dblFactor = FactorDiffusionCoeff(eNewUnits)
    If dblFactor = -1 Or dblFactor = 0 Then
        ConvertDiffusionCoefficient = -1
    Else
        ConvertDiffusionCoefficient = dblValue / dblFactor
    End If

End Function

Public Function ConvertLength(dblLengthIn As Double, eCurrentUnits As ulnUnitsLengthConstants, eNewUnits As ulnUnitsLengthConstants) As Double
    Dim dblValue As Double, dblFactor As Double
    
    If eCurrentUnits = eNewUnits Then
        ConvertLength = dblLengthIn
        Exit Function
    End If
    
    dblFactor = FactorLength(eCurrentUnits)
    If dblFactor = -1 Then
        ConvertLength = -1
        Exit Function
    Else
        dblValue = dblLengthIn * dblFactor
    End If
    
    dblFactor = FactorLength(eNewUnits)
    If dblFactor = -1 Or dblFactor = 0 Then
        ConvertLength = -1
    Else
        ConvertLength = dblValue / dblFactor
    End If

End Function

Public Function ConvertLinearVelocity(dblLinearVelocityIn As Double, eCurrentUnits As ulvUnitsLinearVelocityConstants, eNewUnits As ulvUnitsLinearVelocityConstants) As Double
    Dim dblValue As Double, dblFactor As Double
    
    If eCurrentUnits = eNewUnits Then
        ConvertLinearVelocity = dblLinearVelocityIn
        Exit Function
    End If
    
    dblFactor = FactorLinearVelocity(eCurrentUnits)
    If dblFactor = -1 Then
        ConvertLinearVelocity = -1
        Exit Function
    Else
        dblValue = dblLinearVelocityIn * dblFactor
    End If
    
    dblFactor = FactorLinearVelocity(eNewUnits)
    If dblFactor = -1 Or dblFactor = 0 Then
        ConvertLinearVelocity = -1
    Else
        ConvertLinearVelocity = dblValue / dblFactor
    End If

End Function

Public Function ConvertMassFlowRate(dblMassFlowRateIn As Double, eCurrentUnits As umfMassFlowRateConstants, eNewUnits As umfMassFlowRateConstants) As Double
    Dim dblValue As Double, dblFactor As Double
    
    If eCurrentUnits = eNewUnits Then
        ConvertMassFlowRate = dblMassFlowRateIn
        Exit Function
    End If
    
    dblFactor = FactorMassFlowRate(eCurrentUnits)
    If dblFactor = -1 Then
        ConvertMassFlowRate = -1
        Exit Function
    Else
        dblValue = dblMassFlowRateIn * dblFactor
    End If
    
    dblFactor = FactorMassFlowRate(eNewUnits)
    If dblFactor = -1 Or dblFactor = 0 Then
        ConvertMassFlowRate = -1
    Else
        ConvertMassFlowRate = dblValue / dblFactor
    End If

End Function

Public Function ConvertMoles(dblMolesIn As Double, eCurrentUnits As umaMolarAmountConstants, eNewUnits As umaMolarAmountConstants) As Double
    Dim dblValue As Double, dblFactor As Double
    
    If eCurrentUnits = eNewUnits Then
        ConvertMoles = dblMolesIn
        Exit Function
    End If
    
    dblFactor = FactorMoles(eCurrentUnits)
    If dblFactor = -1 Then
        ConvertMoles = -1
        Exit Function
    Else
        dblValue = dblMolesIn * dblFactor
    End If
    
    dblFactor = FactorMoles(eNewUnits)
    If dblFactor = -1 Or dblFactor = 0 Then
        ConvertMoles = -1
    Else
        ConvertMoles = dblValue / dblFactor
    End If

End Function


Public Function ConvertPressure(dblPressureIn As Double, eCurrentUnits As uprUnitsPressureConstants, eNewUnits As uprUnitsPressureConstants) As Double
    Dim dblValue As Double, dblFactor As Double
    
    If eCurrentUnits = eNewUnits Then
        ConvertPressure = dblPressureIn
        Exit Function
    End If
    
    dblFactor = FactorPressure(eCurrentUnits)
    If dblFactor = -1 Then
        ConvertPressure = -1
        Exit Function
    Else
        dblValue = dblPressureIn * dblFactor
    End If
    
    dblFactor = FactorPressure(eNewUnits)
    If dblFactor = -1 Or dblFactor = 0 Then
        ConvertPressure = -1
    Else
        ConvertPressure = dblValue / dblFactor
    End If

End Function

Public Function ConvertTemperature(dblTemperatureIn As Double, eCurrentUnits As utpUnitsTemperatureConstants, eNewUnits As utpUnitsTemperatureConstants) As Double
    Dim dblValue As Double
    
    If eCurrentUnits = eNewUnits Then
        ConvertTemperature = dblTemperatureIn
        Exit Function
    End If
    
    ' First convert to Kelvin
    Select Case eCurrentUnits
    Case utpCelsius
        ' K = C + 273
        dblValue = dblTemperatureIn + 273
    Case utpFahrenheit
        ' Convert to Kelvin: C = 5/9*(F-32) and K = C + 273
        dblValue = 5# / 9# * (dblTemperatureIn - 32) + 273
    Case Else
        ' Includes utpKelvin
        ' Assume already Kelvin
    End Select
    
    ' We cannot get colder than absolute 0
    If dblValue < 0 Then dblValue = 0
    
    ' Now convert to the target units
    Select Case eNewUnits
    Case utpCelsius
        ' C = K - 273
        dblValue = dblValue - 273
    Case utpFahrenheit
        ' Convert to Fahrenheit: C = K - 273 and F = (9/5)C + 32
        dblValue = 9# / 5# * (dblValue - 273) + 32
    Case Else
        ' Includes utpKelvin
        ' Already in Kelvin
    End Select
    
    ConvertTemperature = dblValue

End Function

Public Function ConvertTime(dblTimeIn As Double, eCurrentUnits As utmUnitsTimeConstants, eNewUnits As utmUnitsTimeConstants) As Double
    Dim dblValue As Double, dblFactor As Double
    
    If eCurrentUnits = eNewUnits Then
        ConvertTime = dblTimeIn
        Exit Function
    End If
    
    dblFactor = FactorTime(eCurrentUnits)
    If dblFactor = -1 Then
        ConvertTime = -1
        Exit Function
    Else
        dblValue = dblTimeIn * dblFactor
    End If
    
    dblFactor = FactorTime(eNewUnits)
    If dblFactor = -1 Or dblFactor = 0 Then
        ConvertTime = -1
    Else
        ConvertTime = dblValue / dblFactor
    End If

End Function

Public Function ConvertViscosity(dblViscosityIn As Double, eCurrentUnits As uviUnitsViscosityConstants, eNewUnits As uviUnitsViscosityConstants) As Double
    Dim dblValue As Double, dblFactor As Double
    
    If eCurrentUnits = eNewUnits Then
        ConvertViscosity = dblViscosityIn
        Exit Function
    End If
    
    dblFactor = FactorViscosity(eCurrentUnits)
    If dblFactor = -1 Then
        ConvertViscosity = -1
        Exit Function
    Else
        dblValue = dblViscosityIn * dblFactor
    End If
    
    dblFactor = FactorViscosity(eNewUnits)
    If dblFactor = -1 Or dblFactor = 0 Then
        ConvertViscosity = -1
    Else
        ConvertViscosity = dblValue / dblFactor
    End If

End Function

Public Function ConvertVolFlowRate(dblVolFlowRateIn As Double, eCurrentUnits As ufrUnitsFlowRateConstants, eNewUnits As ufrUnitsFlowRateConstants) As Double
    Dim dblValue As Double, dblFactor As Double
    
    If eCurrentUnits = eNewUnits Then
        ConvertVolFlowRate = dblVolFlowRateIn
        Exit Function
    End If
    
    dblFactor = FactorVolFlowRate(eCurrentUnits)
    If dblFactor = -1 Then
        ConvertVolFlowRate = -1
        Exit Function
    Else
        dblValue = dblVolFlowRateIn * dblFactor
    End If
    
    dblFactor = FactorVolFlowRate(eNewUnits)
    If dblFactor = -1 Or dblFactor = 0 Then
        ConvertVolFlowRate = -1
    Else
        ConvertVolFlowRate = dblValue / dblFactor
    End If

End Function

Public Function ConvertVolume(dblVolume As Double, eCurrentUnits As uvoUnitsVolumeConstants, eNewUnits As uvoUnitsVolumeConstants) As Double
    Dim dblValue As Double, dblFactor As Double
    
    If eCurrentUnits = eNewUnits Then
        ConvertVolume = dblVolume
        Exit Function
    End If
    
    dblFactor = FactorVolume(eCurrentUnits)
    If dblFactor = -1 Then
        ConvertVolume = -1
        Exit Function
    Else
        dblValue = dblVolume * dblFactor
    End If
    
    dblFactor = FactorVolume(eNewUnits)
    If dblFactor = -1 Or dblFactor = 0 Then
        ConvertVolume = -1
    Else
        ConvertVolume = dblValue / dblFactor
    End If

End Function

' Multiplication factor for converting from eUnits to M
' dblSampleMass is required for mass-based units
' Duplicated function, in both MWCapillaryFlowClass and MWMoleMassDilutionClass
Private Function FactorConcentration(eUnits As ucoUnitsConcentrationConstants, Optional ByVal dblSampleMass As Double = 0) As Double
    Dim dblFactor As Double
    
    If dblSampleMass = 0 Then
        dblFactor = -1
    Else
        Select Case eUnits
        Case ucoMolar:      dblFactor = 1#
        Case ucoMilliMolar: dblFactor = 1 / 1000#
        Case ucoMicroMolar: dblFactor = 1 / 1000000#
        Case ucoNanoMolar:  dblFactor = 1 / 1000000000#
        Case ucoPicoMolar:  dblFactor = 1 / 1000000000000#
        Case ucoFemtoMolar: dblFactor = 1 / 1E+15
        Case ucoAttoMolar:  dblFactor = 1 / 1E+18
        Case ucoMgPerML:    dblFactor = 1 / dblSampleMass           '1/[(1 g / 1000 mg) * (1 / MW) * (1000 mL/L)]
        Case ucoUgPerML:    dblFactor = 1 / (dblSampleMass * 1000#)   '1/[(1 g / 1000000 ug) * (1 / MW) * (1000 mL/L)]
        Case ucongperml:    dblFactor = 1 / (dblSampleMass * 1000000#) '1/[(1 g / 1000000000 ng) * (1 / MW) * (1000 mL/L)]
        Case ucoUgPerUL:    dblFactor = 1 / (dblSampleMass)           '1/[(1 g / 1000000 ug) * (1 / MW) * (1000000 uL/L)]
        Case ucoNgPerUL:    dblFactor = 1 / (dblSampleMass * 1000#)   '1/[(1 g / 1000000000 ng) * (1 / MW) * (1000000 uL/L)]
        Case Else:          dblFactor = -1
        End Select
    End If
    
    FactorConcentration = dblFactor
End Function

' Multiplication factor for converting from eUnits to Cm
Private Function FactorLength(eUnits As ulnUnitsLengthConstants) As Double
    
    Select Case eUnits
    Case ulnM:       FactorLength = 100#
    Case ulnCM:      FactorLength = 1#
    Case ulnMM:      FactorLength = 1 / 10#
    Case ulnMicrons: FactorLength = 1 / 10000#
    Case ulnInches:  FactorLength = CM_PER_INCH
    Case Else:       FactorLength = -1
    End Select

End Function

' Multiplication factor for converting from eUnits to Cm/Min
Private Function FactorLinearVelocity(eUnits As ulvUnitsLinearVelocityConstants) As Double
    
    Select Case eUnits
    Case ulvCmPerHr:  FactorLinearVelocity = 1 / 60#
    Case ulvMmPerHr:  FactorLinearVelocity = 1 / 60# / 10#
    Case ulvCmPerMin: FactorLinearVelocity = 1
    Case ulvMmPerMin: FactorLinearVelocity = 1 / 10#
    Case ulvCmPerSec: FactorLinearVelocity = 60#
    Case ulvMmPerSec: FactorLinearVelocity = 60# / 10#
    Case Else:        FactorLinearVelocity = -1
    End Select

End Function

' Multiplication factor for converting from eUnits to moles/min
Private Function FactorMassFlowRate(eUnits As umfMassFlowRateConstants) As Double
    
    Select Case eUnits
    Case umfPmolPerMin:    FactorMassFlowRate = 1 / 1000000000000#
    Case umfFmolPerMin:    FactorMassFlowRate = 1 / 1E+15
    Case umfAmolPerMin:    FactorMassFlowRate = 1 / 1E+18
    Case umfPmolPerSec:    FactorMassFlowRate = 1 / (1000000000000# / 60#)
    Case umfFmolPerSec:    FactorMassFlowRate = 1 / (1E+15 / 60#)
    Case umfAmolPerSec:    FactorMassFlowRate = 1 / (1E+18 / 60#)
    Case umfMolesPerMin:   FactorMassFlowRate = 1#
    Case Else: FactorMassFlowRate = -1
    End Select

End Function

' Multiplication factor for converting from eUnits to moles
Private Function FactorMoles(eUnits As umaMolarAmountConstants) As Double
    
    Select Case eUnits
    Case umaMoles:      FactorMoles = 1#
    Case umaMilliMoles: FactorMoles = 1 / 1000#
    Case umaMicroMoles: FactorMoles = 1 / 1000000#
    Case umaNanoMoles:  FactorMoles = 1 / 1000000000#
    Case umaPicoMoles:  FactorMoles = 1 / 1000000000000#
    Case umaFemtoMoles: FactorMoles = 1 / 1E+15
    Case umaAttoMoles:  FactorMoles = 1 / 1E+18
    Case Else:          FactorMoles = -1
    End Select
    
End Function

' Multiplication factor for converting from eUnits to dynes per cm^2
Private Function FactorPressure(eUnits As uprUnitsPressureConstants) As Double

    Select Case eUnits
    Case uprPsi:                FactorPressure = 68947.57
    Case uprPascals:            FactorPressure = 10#
    Case uprKiloPascals:        FactorPressure = 10000#
    Case uprAtmospheres:        FactorPressure = 1013250#
    Case uprBar:                FactorPressure = 1000000#
    Case uprTorr:               FactorPressure = 1333.22
    Case uprDynesPerSquareCm:   FactorPressure = 1
    Case Else:                  FactorPressure = -1
    End Select

End Function

' Multiplication factor for converting from eUnits to minutes
Private Function FactorTime(eUnits As utmUnitsTimeConstants) As Double
    
    Select Case eUnits
    Case utmHours:      FactorTime = 60#
    Case utmMinutes:    FactorTime = 1#
    Case utmSeconds:    FactorTime = 1 / 60#
    Case Else:          FactorTime = -1
    End Select

End Function

' Multiplication factor for converting from eUnits to cm^2/sec
Private Function FactorDiffusionCoeff(eUnits As udcDiffusionCoefficientConstants) As Double
    
    Select Case eUnits
    Case udcCmSquaredPerHr:     FactorDiffusionCoeff = 1 / 3600#
    Case udcCmSquaredPerMin:    FactorDiffusionCoeff = 1 / 60#
    Case udcCmSquaredPerSec:    FactorDiffusionCoeff = 1#
    Case Else:                  FactorDiffusionCoeff = -1
    End Select

End Function

' Multiplication factor for converting from eUnits to poise
Private Function FactorViscosity(eUnits As uviUnitsViscosityConstants) As Double
    
    Select Case eUnits
    Case uviPoise:  FactorViscosity = 1#
    Case uviCentiPoise:  FactorViscosity = 1 / 100#
    Case Else:      FactorViscosity = -1
    End Select

End Function

' Multiplication factor for converting from eUnits to mL/min
Private Function FactorVolFlowRate(eUnits As ufrUnitsFlowRateConstants) As Double
    
    Select Case eUnits
    Case ufrMLPerMin:   FactorVolFlowRate = 1#
    Case ufrULPerMin:   FactorVolFlowRate = 1 / 1000#
    Case ufrNLPerMin:   FactorVolFlowRate = 1 / 1000000#
    Case Else:          FactorVolFlowRate = -1
    End Select

End Function

' Multiplication factor for converting from eUnits to mL
Private Function FactorVolume(eUnits As uvoUnitsVolumeConstants) As Double
    
    Select Case eUnits
    Case uvoML:     FactorVolume = 1#
    Case uvoUL:     FactorVolume = 1 / 1000#
    Case uvoNL:     FactorVolume = 1 / 1000000#
    Case uvoPL:     FactorVolume = 1 / 1000000000#
    Case Else:      FactorVolume = -1
    End Select

End Function

' Get Subs
' Gets the most recently computed value
' If mAutoCompute = False, then must manually call a Compute Sub to recompute the value

Public Function GetAutoComputeEnabled() As Boolean
    GetAutoComputeEnabled = mAutoCompute
End Function

Public Function GetAutoComputeMode() As acmAutoComputeModeConstants
    GetAutoComputeMode = mAutoComputeMode
End Function

Public Function GetBackPressure(Optional eUnits As uprUnitsPressureConstants = uprPsi) As Double
    GetBackPressure = ConvertPressure(mCapillaryFlowParameters.BackPressure, uprDynesPerSquareCm, eUnits)
End Function

Public Function GetCapillaryType() As ctCapillaryTypeConstants
    GetCapillaryType = mCapillaryFlowParameters.CapillaryType
End Function

Public Function GetColumnID(Optional eUnits As ulnUnitsLengthConstants = ulnMicrons) As Double
    GetColumnID = ConvertLength(mCapillaryFlowParameters.ColumnID, ulnCM, eUnits)
End Function

Public Function GetColumnLength(Optional eUnits As ulnUnitsLengthConstants = ulnCM) As Double
    GetColumnLength = ConvertLength(mCapillaryFlowParameters.ColumnLength, ulnCM, eUnits)
End Function

Public Function GetColumnVolume(Optional eUnits As uvoUnitsVolumeConstants = uvoUL) As Double
    ' Column volume isn't stored; simply re-compute it
    GetColumnVolume = ComputeColumnVolume(eUnits)
End Function

Public Function GetDeadTime(Optional eUnits As utmUnitsTimeConstants = utmMinutes) As Double
    GetDeadTime = ConvertTime(mCapillaryFlowParameters.ColumnDeadTime, utmMinutes, eUnits)
End Function

Public Function GetExtraColumnBroadeningAdditionalVarianceInSquareSeconds() As Double
    GetExtraColumnBroadeningAdditionalVarianceInSquareSeconds = mExtraColumnBroadeningParameters.AdditionalTemporalVariance
End Function

Public Function GetExtraColumnBroadeningDiffusionCoefficient(Optional eUnits As udcDiffusionCoefficientConstants = udcCmSquaredPerSec) As Double
    GetExtraColumnBroadeningDiffusionCoefficient = ConvertDiffusionCoefficient(mExtraColumnBroadeningParameters.DiffusionCoefficient, udcCmSquaredPerSec, eUnits)
End Function

Public Function GetExtraColumnBroadeningInitialPeakWidthAtBase(Optional eUnits As utmUnitsTimeConstants = utmSeconds) As Double
    GetExtraColumnBroadeningInitialPeakWidthAtBase = ConvertTime(mExtraColumnBroadeningParameters.InitialPeakWidth, utmSeconds, eUnits)
End Function

Public Function GetExtraColumnBroadeningLinearVelocity(Optional eUnits As ulvUnitsLinearVelocityConstants = ulvMmPerMin) As Double
    GetExtraColumnBroadeningLinearVelocity = ConvertLinearVelocity(mExtraColumnBroadeningParameters.LinearVelocity, ulvCmPerMin, eUnits)
End Function

Public Function GetExtraColumnBroadeningOpenTubeID(Optional eUnits As ulnUnitsLengthConstants = ulnMicrons) As Double
    GetExtraColumnBroadeningOpenTubeID = ConvertLength(mExtraColumnBroadeningParameters.OpenTubeID, ulnCM, eUnits)
End Function

Public Function GetExtraColumnBroadeningOpenTubeLength(Optional eUnits As ulnUnitsLengthConstants = ulnCM) As Double
   GetExtraColumnBroadeningOpenTubeLength = ConvertLength(mExtraColumnBroadeningParameters.OpenTubeLength, ulnCM, eUnits)
End Function

Public Function GetExtraColumnBroadeningResultantPeakWidth(Optional eUnits As utmUnitsTimeConstants = utmSeconds) As Double
    GetExtraColumnBroadeningResultantPeakWidth = ConvertTime(mExtraColumnBroadeningParameters.ResultantPeakWidth, utmSeconds, eUnits)
End Function

Public Function GetExtraColumnBroadeningTemporalVarianceInSquareSeconds() As Double
    GetExtraColumnBroadeningTemporalVarianceInSquareSeconds = mExtraColumnBroadeningParameters.TemporalVariance
End Function

Public Function GetInterparticlePorosity() As Double
    GetInterparticlePorosity = mCapillaryFlowParameters.InterparticlePorosity
End Function

Public Function GetLinearVelocity(Optional eUnits As ulvUnitsLinearVelocityConstants = ulvCmPerSec) As Double
    GetLinearVelocity = ConvertLinearVelocity(mCapillaryFlowParameters.LinearVelocity, ulvCmPerMin, eUnits)
End Function

Public Function GetMassRateConcentration(Optional eUnits As ucoUnitsConcentrationConstants = ucoMicroMolar) As Double
    GetMassRateConcentration = ConvertConcentration(mMassRateParameters.SampleConcentration, ucoMolar, eUnits)
End Function

Public Function GetMassRateInjectionTime(Optional eUnits As utmUnitsTimeConstants = utmMinutes) As Double
    GetMassRateInjectionTime = ConvertTime(mMassRateParameters.InjectionTime, utmMinutes, eUnits)
End Function

Public Function GetMassFlowRate(Optional eUnits As umfMassFlowRateConstants = umfFmolPerSec) As Double
    GetMassFlowRate = ConvertMassFlowRate(mMassRateParameters.MassFlowRate, umfMolesPerMin, eUnits)
End Function

Public Function GetMassRateMolesInjected(Optional eUnits As umaMolarAmountConstants = umaFemtoMoles) As Double
    GetMassRateMolesInjected = ConvertMoles(mMassRateParameters.MolesInjected, umaMoles, eUnits)
End Function

Public Function GetMassRateSampleMass() As Double
    GetMassRateSampleMass = mMassRateParameters.SampleMass
End Function

Public Function GetMassRateVolFlowRate(Optional eUnits As ufrUnitsFlowRateConstants = ufrNLPerMin) As Double
    GetMassRateVolFlowRate = ConvertVolFlowRate(mMassRateParameters.VolumetricFlowRate, ufrMLPerMin, eUnits)
End Function

Public Function GetParticleDiameter(Optional eUnits As ulnUnitsLengthConstants = ulnMicrons) As Double
    GetParticleDiameter = ConvertLength(mCapillaryFlowParameters.ParticleDiameter, ulnCM, eUnits)
End Function

Public Function GetSolventViscosity(Optional eUnits As uviUnitsViscosityConstants = uviPoise) As Double
    GetSolventViscosity = ConvertViscosity(mCapillaryFlowParameters.SolventViscosity, uviPoise, eUnits)
End Function

Public Function GetVolFlowRate(Optional eUnits As ufrUnitsFlowRateConstants = ufrNLPerMin) As Double
    GetVolFlowRate = ConvertVolFlowRate(mCapillaryFlowParameters.VolumetricFlowRate, ufrMLPerMin, eUnits)
End Function


' Set Subs
' If mAutoCompute = False, then must manually call a Compute Sub to recompute other values

Public Sub SetAutoComputeEnabled(blnAutoCompute As Boolean)
    mAutoCompute = blnAutoCompute
End Sub

Public Sub SetAutoComputeMode(eAutoComputeMode As acmAutoComputeModeConstants)
    If eAutoComputeMode >= acmBackPressure And eAutoComputeMode <= acmVolFlowrateUsingDeadTime Then
        mAutoComputeMode = eAutoComputeMode
    End If
End Sub

Public Sub SetBackPressure(dblBackPressure As Double, Optional eUnits As uprUnitsPressureConstants = uprPsi)
    mCapillaryFlowParameters.BackPressure = ConvertPressure(dblBackPressure, eUnits, uprDynesPerSquareCm)
    CheckAutoCompute
End Sub

Public Sub SetCapillaryType(eCapillaryType As ctCapillaryTypeConstants)
    If eCapillaryType >= ctOpenTubularCapillary And eCapillaryType <= ctPackedCapillary Then
        mCapillaryFlowParameters.CapillaryType = eCapillaryType
    End If
    CheckAutoCompute
End Sub

Public Sub SetColumnID(dblColumnID As Double, Optional eUnits As ulnUnitsLengthConstants = ulnMicrons)
    mCapillaryFlowParameters.ColumnID = ConvertLength(dblColumnID, eUnits, ulnCM)
    CheckAutoCompute
End Sub

Public Sub SetColumnLength(dblColumnLength As Double, Optional eUnits As ulnUnitsLengthConstants = ulnCM)
    mCapillaryFlowParameters.ColumnLength = ConvertLength(dblColumnLength, eUnits, ulnCM)
    CheckAutoCompute
End Sub

Public Sub SetDeadTime(dblDeadTime As Double, Optional eUnits As utmUnitsTimeConstants = utmMinutes)
    mCapillaryFlowParameters.ColumnDeadTime = ConvertTime(dblDeadTime, eUnits, utmMinutes)
    CheckAutoCompute
End Sub

Public Sub SetExtraColumnBroadeningAdditionalVariance(dblAdditionalVarianceInSquareSeconds As Double)
    mExtraColumnBroadeningParameters.AdditionalTemporalVariance = dblAdditionalVarianceInSquareSeconds
    ComputeExtraColumnBroadeningValues
End Sub

Public Sub SetExtraColumnBroadeningDiffusionCoefficient(dblDiffusionCoefficient As Double, Optional eUnits As udcDiffusionCoefficientConstants = udcCmSquaredPerSec)
    mExtraColumnBroadeningParameters.DiffusionCoefficient = ConvertDiffusionCoefficient(dblDiffusionCoefficient, eUnits, udcCmSquaredPerSec)
    ComputeExtraColumnBroadeningValues
End Sub

Public Sub SetExtraColumnBroadeningInitialPeakWidthAtBase(dblWidth As Double, Optional eUnits As utmUnitsTimeConstants = utmSeconds)
    mExtraColumnBroadeningParameters.InitialPeakWidth = ConvertTime(dblWidth, eUnits, utmSeconds)
    ComputeExtraColumnBroadeningValues
End Sub

Public Sub SetExtraColumnBroadeningLinearVelocity(dblLinearVelocity As Double, Optional eUnits As ulvUnitsLinearVelocityConstants = ulvMmPerMin)
    mExtraColumnBroadeningParameters.LinearVelocity = ConvertLinearVelocity(dblLinearVelocity, eUnits, ulvCmPerMin)
    ComputeExtraColumnBroadeningValues
End Sub

Public Sub SetExtraColumnBroadeningOpenTubeID(dblOpenTubeID As Double, Optional eUnits As ulnUnitsLengthConstants = ulnMicrons)
    mExtraColumnBroadeningParameters.OpenTubeID = ConvertLength(dblOpenTubeID, eUnits, ulnCM)
    ComputeExtraColumnBroadeningValues
End Sub

Public Sub SetExtraColumnBroadeningOpenTubeLength(dblLength As Double, Optional eUnits As ulnUnitsLengthConstants = ulnCM)
    mExtraColumnBroadeningParameters.OpenTubeLength = ConvertLength(dblLength, eUnits, ulnCM)
    ComputeExtraColumnBroadeningValues
End Sub

Public Sub SetInterparticlePorosity(dblPorosity As Double)
    If dblPorosity >= 0 And dblPorosity <= 1 Then
        mCapillaryFlowParameters.InterparticlePorosity = dblPorosity
    End If
    CheckAutoCompute
End Sub

Public Sub SetMassRateConcentration(dblConcentration As Double, Optional eUnits As ucoUnitsConcentrationConstants = ucoMicroMolar)
    mMassRateParameters.SampleConcentration = ConvertConcentration(dblConcentration, eUnits, ucoMolar)
    ComputeMassRateValues
End Sub

Public Sub SetMassRateInjectionTime(dblInjectionTime As Double, Optional eUnits As utmUnitsTimeConstants = utmMinutes)
    mMassRateParameters.InjectionTime = ConvertTime(dblInjectionTime, eUnits, utmMinutes)
    ComputeMassRateValues
End Sub

Public Function SetMassRateSampleMass(dblMassInGramsPerMole As Double) As Double
    If dblMassInGramsPerMole >= 0 Then
        mMassRateParameters.SampleMass = dblMassInGramsPerMole
    Else
        mMassRateParameters.SampleMass = 0
    End If
    ComputeMassRateValues
End Function

Public Sub SetMassRateVolFlowRate(dblVolFlowRate As Double, Optional eUnits As ufrUnitsFlowRateConstants = ufrNLPerMin)
    mMassRateParameters.VolumetricFlowRate = ConvertVolFlowRate(dblVolFlowRate, eUnits, ufrMLPerMin)
    ComputeMassRateValues
End Sub

Public Sub SetParticleDiameter(dblParticleDiameter As Double, Optional eUnits As ulnUnitsLengthConstants = ulnMicrons)
    mCapillaryFlowParameters.ParticleDiameter = ConvertLength(dblParticleDiameter, eUnits, ulnCM)
    CheckAutoCompute
End Sub

Public Sub SetSolventViscosity(dblSolventViscosity As Double, Optional eUnits As uviUnitsViscosityConstants = uviPoise)
    mCapillaryFlowParameters.SolventViscosity = ConvertViscosity(dblSolventViscosity, eUnits, uviPoise)
    CheckAutoCompute
End Sub

Public Sub SetVolFlowRate(dblVolFlowRate As Double, Optional eUnits As ufrUnitsFlowRateConstants = ufrNLPerMin)
    mCapillaryFlowParameters.VolumetricFlowRate = ConvertVolFlowRate(dblVolFlowRate, eUnits, ufrMLPerMin)
    CheckAutoCompute
End Sub

Private Sub Class_Initialize()
    Me.SetAutoComputeEnabled False
    
    Me.SetAutoComputeMode acmVolFlowrate
    Me.SetCapillaryType ctPackedCapillary
    Me.SetBackPressure 3000, uprPsi
    Me.SetColumnLength 50, ulnCM
    Me.SetColumnID 75, ulnMicrons
    Me.SetSolventViscosity 0.0089, uviPoise
    Me.SetParticleDiameter 5, ulnMicrons
    Me.SetInterparticlePorosity 0.4

    Me.SetMassRateConcentration 1, ucoMicroMolar
    Me.SetMassRateVolFlowRate 600, ufrNLPerMin
    Me.SetMassRateInjectionTime 5, utmMinutes
    
    ' Recompute
    ComputeVolFlowRate
    ComputeMassRateValues
    ComputeExtraColumnBroadeningValues
    
    Me.SetAutoComputeEnabled True

End Sub



