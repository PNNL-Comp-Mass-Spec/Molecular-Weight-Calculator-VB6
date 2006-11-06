VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmMwtWinDllTest 
   Caption         =   "Mwt Win Dll Test"
   ClientHeight    =   6690
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   9780
   LinkTopic       =   "Form1"
   ScaleHeight     =   6690
   ScaleWidth      =   9780
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTestGetTrypticName 
      Caption         =   "Test Get Tryptic Name"
      Height          =   615
      Left            =   3840
      TabIndex        =   19
      Top             =   4440
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid grdFlexGrid 
      Height          =   5000
      Left            =   5280
      TabIndex        =   7
      Top             =   960
      Width           =   4000
      _ExtentX        =   7064
      _ExtentY        =   8811
      _Version        =   393216
   End
   Begin VB.CommandButton cmdExpandAbbreviations 
      Caption         =   "Expand Abbreviations"
      Height          =   615
      Left            =   4320
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdTestFunctions 
      Caption         =   "Test Functions"
      Height          =   615
      Left            =   7680
      TabIndex        =   6
      Top             =   120
      Width           =   1335
   End
   Begin VB.ComboBox cboStdDevMode 
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   3720
      Width           =   2175
   End
   Begin VB.ComboBox cboWeightMode 
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   3360
      Width           =   2175
   End
   Begin VB.CommandButton cmdConvertToEmpirical 
      Caption         =   "Convert to &Empirical"
      Height          =   615
      Left            =   6000
      TabIndex        =   5
      Top             =   120
      Width           =   1335
   End
   Begin RichTextLib.RichTextBox rtfFormula 
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   1320
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   873
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"MwtWinDllTest.frx":0000
   End
   Begin VB.CommandButton cmdFindMass 
      Caption         =   "&Calculate"
      Default         =   -1  'True
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   4560
      Width           =   1335
   End
   Begin VB.TextBox txtFormula 
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Text            =   "Cl2PhH4OH"
      Top             =   240
      Width           =   1935
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Cl&ose"
      Height          =   495
      Left            =   1560
      TabIndex        =   9
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Label lblStdDevMode 
      Caption         =   "Std Dev Mode"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label lblWeightMode 
      Caption         =   "Weight Mode"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label lblStatusLabel 
      Caption         =   "Status:"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label lblMassAndStdDevLabel 
      Caption         =   "Mass and StdDev:"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label lblMassAndStdDev 
      Caption         =   "0"
      Height          =   255
      Left            =   1560
      TabIndex        =   14
      Top             =   960
      Width           =   3135
   End
   Begin VB.Label lblStatus 
      Height          =   1095
      Left            =   1560
      TabIndex        =   13
      Top             =   1920
      Width           =   3375
   End
   Begin VB.Label lblMass 
      Caption         =   "0"
      Height          =   255
      Left            =   1560
      TabIndex        =   12
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label lblMassLabel 
      Caption         =   "Mass:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label lblFormula 
      Caption         =   "Formula:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "frmMwtWinDllTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Molecular Weight Calculator Dll test program
' Written by Matthew Monroe for the Department of Energy (PNNL, Richland, WA) in Richland, WA
' Copyright 2005, Battelle Memorial Institute

Private Sub FindPercentComposition()
    Dim intIndex As Integer, intElementCount As Integer
    Dim dblPctCompForCarbon As Double, strPctCompForCarbon As String
    Dim strSymbol As String
    
    Dim strPctCompositions() As String
    
    With objMwtWin.Compound
        .Formula = txtFormula
        
        dblPctCompForCarbon = .GetPercentCompositionForElement(6)
        strPctCompForCarbon = .GetPercentCompositionForElementAsString(6)
        
        intElementCount = objMwtWin.GetElementCount
        ReDim strPctCompositions(intElementCount)
        
        .GetPercentCompositionForAllElements strPctCompositions()
        
    End With
    
    With grdFlexGrid
        .Clear
        .Rows = 2
        .Cols = 2
        .ColWidth(1) = 1500
        .FixedCols = 0
        .FixedRows = 1
        .TextMatrix(0, 0) = "Element": .TextMatrix(0, 1) = "Pct Comp"
        
        For intIndex = 1 To intElementCount
            If Left(strPctCompositions(intIndex), 5) <> "" Then
                objMwtWin.GetElement intIndex, strSymbol, 0, 0, 0, 0
                .TextMatrix(.Rows - 1, 0) = strSymbol
                .TextMatrix(.Rows - 1, 1) = strPctCompositions(intIndex)
                .Rows = .Rows + 1
            End If
        Next intIndex
    End With
End Sub

Private Sub FindMass()
    ' Can simply compute the mass of a formula using ComputeMass
    lblMass = objMwtWin.ComputeMass(txtFormula)
    
    ' If we want to do more complex operations, need to fill objMwtWin.Compound with valid info
    ' Then, can read out values from it
    With objMwtWin.Compound
        .Formula = txtFormula

        If .ErrorDescription = "" Then
            lblMass = .Mass
            lblStatus = .CautionDescription
            txtFormula = .FormulaCapitalized
            rtfFormula.TextRTF = .FormulaRTF
            lblMassAndStdDev = .MassAndStdDevString
        Else
            lblStatus = .ErrorDescription
        End If
    End With

End Sub

Private Sub PopulateComboBoxes()
    With cboWeightMode
        .Clear
        .AddItem "Average mass"
        .AddItem "Isotopic mass"
        .AddItem "Integer mass"
        .ListIndex = 0
    End With

    With cboStdDevMode
        .Clear
        .AddItem "Short"
        .AddItem "Scientific"
        .AddItem "Decimal"
        .ListIndex = 0
    End With


End Sub

Public Sub TestAccessFunctions()
    Dim intIndex As Integer, intResult As Integer, lngIndex As Long
    Dim lngItemCount As Long
    Dim strSymbol As String, strFormula As String, sngCharge As Single, blnIsAminoAcid As Boolean
    Dim strOneLetterSymbol As String, strComment As String
    Dim strStatement As String
    Dim dblmass As Double, dblUncertainty As Double, intIsotopeCount As Integer, intIsotopeCount2 As Integer
    Dim dblIsotopeMasses() As Double, sngIsotopeAbundances() As Single
    Dim dblNewPressure As Double
    
    Dim objCompound As New MWCompoundClass
    
    With objMwtWin
        ' Test Abbreviations
        lngItemCount = .GetAbbreviationCount
        For intIndex = 1 To lngItemCount
            intResult = .GetAbbreviation(intIndex, strSymbol, strFormula, sngCharge, blnIsAminoAcid, strOneLetterSymbol, strComment)
            Debug.Assert intResult = 0
            Debug.Assert .GetAbbreviationID(strSymbol) = intIndex

            intResult = .SetAbbreviation(strSymbol, strFormula, sngCharge, blnIsAminoAcid, strOneLetterSymbol, strComment)
            Debug.Assert intResult = 0
        Next intIndex
        
        ' Test Caution statements
        lngItemCount = .GetCautionStatementCount
        For intIndex = 1 To lngItemCount
            intResult = .GetCautionStatement(intIndex, strSymbol, strStatement)
            Debug.Assert intResult = 0
            Debug.Assert .GetCautionStatementID(strSymbol) = intIndex
            
            intResult = .SetCautionStatement(strSymbol, strStatement)
            Debug.Assert intResult = 0
        Next intIndex
        
        ' Test Element access
        lngItemCount = .GetElementCount
        For intIndex = 1 To lngItemCount
            intResult = .GetElement(intIndex, strSymbol, dblmass, dblUncertainty, sngCharge, intIsotopeCount)
            Debug.Assert intResult = 0
            Debug.Assert .GetElementID(strSymbol) = intIndex
            
            intResult = .SetElement(strSymbol, dblmass, dblUncertainty, sngCharge, False)
            Debug.Assert intResult = 0
            
            ReDim dblIsotopeMasses(intIsotopeCount + 1)
            ReDim sngIsotopeAbundances(intIsotopeCount + 1)
            
            intResult = .GetElementIsotopes(intIndex, intIsotopeCount2, dblIsotopeMasses(), sngIsotopeAbundances())
            Debug.Assert intIsotopeCount = intIsotopeCount2
            Debug.Assert intResult = 0
            
            intResult = .SetElementIsotopes(strSymbol, intIsotopeCount, dblIsotopeMasses(), sngIsotopeAbundances())
            Debug.Assert intResult = 0
        Next intIndex
        
        ' Test Message Statements access
        lngItemCount = .GetMessageStatementCount
        For lngIndex = 1 To lngItemCount
            strStatement = .GetMessageStatement(lngIndex)
            
            intResult = .SetMessageStatement(lngIndex, strStatement)
        Next lngIndex
        
        ' Test Capillary flow functions
        
        With .CapFlow
            .SetAutoComputeEnabled False
            .SetBackPressure 2000, uprPsi
            .SetColumnLength 40, ulnCM
            .SetColumnID 50, ulnMicrons
            .SetSolventViscosity 0.0089, uviPoise
            .SetInterparticlePorosity 0.33
            .SetParticleDiameter 2, ulnMicrons
            .SetAutoComputeEnabled True
            
            Debug.Print ""
            Debug.Print "Check capillary flow calcs"
            Debug.Print "Linear Velocity: " & .ComputeLinearVelocity(ulvCmPerSec)
            Debug.Print "Vol flow rate:   " & .ComputeVolFlowRate(ufrNLPerMin) & "  (newly computed)"
            
            Debug.Print "Vol flow rate:   " & .GetVolFlowRate
            Debug.Print "Back pressure:   " & .ComputeBackPressure(uprPsi)
            Debug.Print "Column Length:   " & .ComputeColumnLength(ulnCM)
            Debug.Print "Column ID:       " & .ComputeColumnID(ulnMicrons)
            Debug.Print "Column Volume:   " & .ComputeColumnVolume(uvoNL)
            Debug.Print "Dead time:       " & .ComputeDeadTime(utmSeconds)
            
            Debug.Print ""
            
            Debug.Print "Repeat Computations, but in a different order (should give same results)"
            Debug.Print "Vol flow rate:   " & .ComputeVolFlowRate(ufrNLPerMin)
            Debug.Print "Column ID:       " & .ComputeColumnID(ulnMicrons)
            Debug.Print "Back pressure:   " & .ComputeBackPressure(uprPsi)
            Debug.Print "Column Length:   " & .ComputeColumnLength(ulnCM)
            
            Debug.Print ""
            
            Debug.Print "Old Dead time: " & .GetDeadTime(utmMinutes)
            
            .SetAutoComputeMode acmVolFlowrateUsingDeadTime
            
            .SetDeadTime 25, utmMinutes
            Debug.Print "Dead time is now 25.0 minutes"
            
            Debug.Print "Vol flow rate: " & .GetVolFlowRate(ufrNLPerMin) & " (auto-computed since AutoComputeMode = acmVolFlowrateUsingDeadTime)"
            
            ' Confirm that auto-compute worked
            
            Debug.Print "Vol flow rate: " & .ComputeVolFlowRateUsingDeadTime(ufrNLPerMin, dblNewPressure, uprPsi) & "  (confirmation of computed volumetric flow rate)"
            Debug.Print "New pressure: " & dblNewPressure
            
            Debug.Print ""
            
            ' Can set a new back pressure, but since auto-compute is on, and the
            '  auto-compute mode is acmVolFlowRateUsingDeadTime, the pressure will get changed back to
            '  the pressure needed to give a vol flow rate matching the dead time
            .SetBackPressure 2000
            Debug.Print "Pressure set to 2000 psi, but auto-compute mode is acmVolFlowRateUsingDeadTime, so pressure"
            Debug.Print "  was automatically changed back to pressure needed to give vol flow rate matching dead time"
            Debug.Print "Pressure is now: " & .GetBackPressure(uprPsi) & " psi (thus, not 2000 as one might expect)"
            
            .SetAutoComputeMode acmVolFlowrate
            Debug.Print "Changed auto-compute mode to acmVolFlowrate.  Can now set pressure to 2000 and it will stick; plus, vol flow rate gets computed."
            
            .SetBackPressure 2000, uprPsi
            
            ' Calling GetVolFlowRate will get the new computed vol flow rate (since auto-compute is on)
            Debug.Print "Vol flow rate: " & .GetVolFlowRate
            
            .SetMassRateConcentration 1, ucoMicroMolar
            .SetMassRateVolFlowRate 600, ufrNLPerMin
            .SetMassRateInjectionTime 5, utmMinutes

            Debug.Print "Mass flow rate: " & .GetMassFlowRate(umfFmolPerSec) & " fmol/sec"
            Debug.Print "Moles injected: " & .GetMassRateMolesInjected(umaFemtoMoles) & " fmoles"

            .SetMassRateSampleMass 1234
            .SetMassRateConcentration 1, ucongperml
            
            Debug.Print "Computing mass flow rate for compound weighing 1234 g/mol and at 1 ng/mL concentration                "
            Debug.Print "Mass flow rate: " & .GetMassFlowRate(umfAmolPerMin) & " amol/min"
            Debug.Print "Moles injected: " & .GetMassRateMolesInjected(umaFemtoMoles) & " fmoles"
            
            .SetExtraColumnBroadeningLinearVelocity 4, ulvCmPerMin
            .SetExtraColumnBroadeningDiffusionCoefficient 0.0003, udcCmSquaredPerMin
            .SetExtraColumnBroadeningOpenTubeLength 5, ulnCM
            .SetExtraColumnBroadeningOpenTubeID 250, ulnMicrons
            .SetExtraColumnBroadeningInitialPeakWidthAtBase 30, utmSeconds
            
            Debug.Print "Computing broadening for 30 second wide peak through a 250 um open tube that is 5 cm long (4 cm/min)"
            Debug.Print .GetExtraColumnBroadeningResultantPeakWidth(utmSeconds)
            
        End With
    End With
    
    
    Dim udtFragSpectrumOptions As udtFragmentationSpectrumOptionsType
    Dim udtFragSpectrum() As udtFragmentationSpectrumDataType
    Dim lngIonCount As Long
    Dim strNewSeq As String
    
    With objMwtWin.Peptide
        
        .SetSequence "K.ACYEFGHRKACYEFGHRK.G", ntgHydrogen, ctgHydroxyl, False, True

        ' Can change the terminii to various standard groups
        .SetNTerminusGroup ntgCarbamyl
        .SetCTerminusGroup ctgAmide
        
        ' Can change the terminii to any desired elements
        .SetNTerminus "C2OH3"       ' Acetyl group
        .SetCTerminus "NH2"         ' Amide group
        
        ' Can mark third residue, Tyr, as phorphorylated
        .SetResidue 3, "Tyr", True, True
        
        ' Can define that the * modification equals 15
        .SetModificationSymbol "*", 15, False, ""
        
        strNewSeq = "Ala-Cys-Tyr-Glu-Phe-Gly-His-Arg*-Lys-Ala-Cys-Tyr-Glu-Phe-Gly-His-Arg-Lys"
        Debug.Print strNewSeq
        .SetSequence strNewSeq
        
        .SetSequence "K.TQPLE*VK.-", ntgHydrogenPlusProton, ctgHydroxyl, False, True
        
        Debug.Print .GetSequence(True, False, True, False)
        Debug.Print .GetSequence(False, True, False, False)
        Debug.Print .GetSequence(True, False, True, True)
        
        .SetCTerminusGroup ctgNone
        Debug.Print .GetSequence(True, False, True, True)
        
        udtFragSpectrumOptions = .GetFragmentationSpectrumOptions()
        
        udtFragSpectrumOptions.DoubleChargeIonsShow = False
        udtFragSpectrumOptions.DoubleChargeIonsThreshold = 200
        udtFragSpectrumOptions.IntensityOptions.BYIonShoulder = 0
        
        udtFragSpectrumOptions.IonTypeOptions(itAIon).ShowIon = True
        
        .SetFragmentationSpectrumOptions udtFragSpectrumOptions
        
        lngIonCount = .GetFragmentationMasses(udtFragSpectrum())
        
    End With
    
    With grdFlexGrid
        .Clear
        
        .Rows = lngIonCount + 1
        .Cols = 3
        .ColWidth(0) = 1500
        .ColWidth(1) = 1000
        .ColWidth(2) = 1000
        
        .FixedCols = 0
        .FixedRows = 1
        .TextMatrix(0, 0) = "Mass"
        .TextMatrix(0, 1) = "Intensity"
        .TextMatrix(0, 2) = "Symbol"
        
        For lngIndex = 0 To lngIonCount - 1
            .TextMatrix(lngIndex + 1, 0) = udtFragSpectrum(lngIndex).Mass
            .TextMatrix(lngIndex + 1, 1) = udtFragSpectrum(lngIndex).Intensity
            .TextMatrix(lngIndex + 1, 2) = udtFragSpectrum(lngIndex).Symbol
            
            If udtFragSpectrum(lngIndex).Symbol <> "" Then
'                Debug.Print sngIonMassesZeroBased(lngIndex) & ", " & sngIonIntensitiesZeroBased(lngIndex) & ", " & strIonSymbolsZeroBased(lngIndex)
            End If
        Next lngIndex
        
        Debug.Print "Peptide mass: " & objMwtWin.Peptide.GetPeptideMass
        
        
    End With
    
        
    Dim intSuccess As Integer
    Dim strResults As String
    Dim ConvolutedMSData2D() As Double, ConvolutedMSDataCount As Long
    
    With objMwtWin
        Debug.Print ""
        intSuccess = .ComputeIsotopicAbundances("C1255H43O2Cl", 1, strResults, ConvolutedMSData2D(), ConvolutedMSDataCount)
        Debug.Print strResults
    End With
    
    
End Sub

Private Sub TestTrypticName()
    Const DIM_CHUNK = 1000
    
    Const ITERATIONS_TO_RUN = 5
    Const MIN_PROTEIN_LENGTH = 50
    Const MAX_PROTEIN_LENGTH = 200
    Const POSSIBLE_RESIDUES = "ACDEFGHIKLMNPQRSTVWY"
    
    Dim lngMultipleIteration As Long
    
    Dim strProtein As String, strPeptideResidues As String
    Dim lngResidueStart As Long, lngResidueEnd As Long
    Dim strPeptideNameMwtWin() As String
    Dim strPeptideNameIcr2ls() As String
    Dim strPeptideName As String
    
    Dim lngMwtWinResultCount As Long, lngIcr2lsResultCount As Long
    Dim lngMwtWinDimCount As Long, lngICR2lsDimCount As Long
    Dim lngIndex As Long
    Dim lngResidueRand As Long, lngProteinLengthRand As Long
    Dim strNewResidue As String
    
    Dim lngStartTime As Long, lngStopTime As Long
    Dim lngMwtWinWorkTime As Long
    Dim strPeptideFragMwtWin As String
    Dim lngMatchCount As Long
    Dim blnDifferenceFound As Boolean
    
    lngMwtWinDimCount = DIM_CHUNK
    ReDim strPeptideNameMwtWin(lngMwtWinDimCount)
    
    Me.MousePointer = vbHourglass
    
    frmResults.Show
    frmResults.txtResults = ""
    
''    Dim lngIcr2lsWorkTime As Long
''    Dim lngIcr2lsTime As Long
''    strPeptideFragIcr2ls As String
''    lngICR2lsDimCount = DIM_CHUNK
''    ReDim strPeptideNameIcr2ls(lngICR2lsDimCount)
''
''    Dim ICRTools As Object
''
''    Set ICRTools = CreateObject("ICR2LS.ICR2LScls")
''
''    frmResults.AddToResults "ICR2ls Version: " & ICRTools.ICR2LSversion

    'strProtein = "MGNISFLTGGNPSSPQSIAESIYQLENTSVVFLSAWQRTTPDFQRAARASQEAMLHLDHIVNEIMRNRDQLQADGTYTGSQLEGLLNISRAVSVSPVTRAEQDDLANYGPGNGVLPSAGSSISMEKLLNKIKHRRTNSANFRIGASGEHIFIIGVDKPNRQPDSIVEFIVGDFCQHCSDIAALI"
    
    ' Bigger protein
    strProtein = "MMKANVTKKTLNEGLGLLERVIPSRSSNPLLTALKVETSEGGLTLSGTNLEIDLSCFVPAEVQQPENFVVPAHLFAQIVRNLGGELVELELSGQELSVRSGGSDFKLQTGDIEAYPPLSFPAQADVSLDGGELSRAFSSVRYAASNEAFQAVFRGIKLEHHGESARVVASDGYRVAIRDFPASGDGKNLIIPARSVDELIRVLKDGEARFTYGDGMLTVTTDRVKMNLKLLDGDFPDYERVIPKDIKLQVTLPATALKEAVNRVAVLADKNANNRVEFLVSEGTLRLAAEGDYGRAQDTLSVTQGGTEQAMSLAFNARHVLDALGPIDGDAELLFSGSTSPAIFRARRWGRRVYGGHGHAARLRGLLRPLRGMSALAHHPESSPPLEPRPEFA"
    
    frmResults.AddToResults "Testing GetTrypticNameMultipleMatches() function"
    frmResults.AddToResults "MatchList for NL: " & objMwtWin.Peptide.GetTrypticNameMultipleMatches(strProtein, "NL", lngMatchCount)
    frmResults.AddToResults "MatchCount = " & lngMatchCount
    
    frmResults.AddToResults ""
    frmResults.AddToResults "Testing GetTrypticPeptideByFragmentNumber function"
    For lngIndex = 1 To 43
        strPeptideFragMwtWin = objMwtWin.Peptide.GetTrypticPeptideByFragmentNumber(strProtein, CInt(lngIndex), lngResidueStart, lngResidueEnd)
''        strPeptideFragIcr2ls = ICRTools.TrypticPeptide(strProtein, CInt(lngIndex))
''
''        Debug.Assert strPeptideFragMwtWin = strPeptideFragIcr2ls
        
        If Len(strPeptideFragMwtWin) > 1 Then
            ' Make sure lngResidueStart and lngResidueEnd are correct
            ' Do this using .GetTrypticNameMultipleMatches()
            strPeptideName = objMwtWin.Peptide.GetTrypticNameMultipleMatches(strProtein, Mid(strProtein, lngResidueStart, lngResidueEnd - lngResidueStart + 1))
            Debug.Assert InStr(strPeptideName, "t" & Trim(Str(lngIndex))) > 0
        End If
    Next lngIndex
    frmResults.AddToResults "Check of GetTrypticPeptideByFragmentNumber Complete"
    frmResults.AddToResults ""
    
    
    frmResults.AddToResults "Test tryptic digest of: " & strProtein
    lngIndex = 1
    Do
        strPeptideFragMwtWin = objMwtWin.Peptide.GetTrypticPeptideByFragmentNumber(strProtein, CInt(lngIndex), lngResidueStart, lngResidueEnd)
        frmResults.AddToResults "Tryptic fragment " & Trim(lngIndex) & ": " & strPeptideFragMwtWin
        lngIndex = lngIndex + 1
    Loop While Len(strPeptideFragMwtWin) > 0
    
    
    frmResults.AddToResults ""
    Randomize
    For lngMultipleIteration = 1 To ITERATIONS_TO_RUN
        ' Generate random protein
        lngProteinLengthRand = Int((MAX_PROTEIN_LENGTH - MIN_PROTEIN_LENGTH + 1) * Rnd + MIN_PROTEIN_LENGTH)

        strProtein = ""
        For lngResidueRand = 1 To lngProteinLengthRand
            strNewResidue = Mid(POSSIBLE_RESIDUES, Int(Len(POSSIBLE_RESIDUES)) * Rnd + 1, 1)
            strProtein = strProtein & strNewResidue
        Next lngResidueRand
        
        frmResults.AddToResults "Iteration: " & lngMultipleIteration & " = " & strProtein
        
        lngMwtWinResultCount = 0
        Debug.Print "Starting residue is ";
        lngStartTime = GetTickCount()
        For lngResidueStart = 1 To Len(strProtein)
            If lngResidueStart Mod 10 = 0 Then
                Debug.Print lngResidueStart & ", ";
                DoEvents
            End If
            
            For lngResidueEnd = 1 To Len(strProtein) - lngResidueStart
                If lngResidueEnd - lngResidueStart > 50 Then
                    Exit For
                End If
                
                strPeptideResidues = Mid(strProtein, lngResidueStart, lngResidueEnd)
                strPeptideNameMwtWin(lngMwtWinResultCount) = objMwtWin.Peptide.GetTrypticName(strProtein, strPeptideResidues, 0, 0, True)
                
                lngMwtWinResultCount = lngMwtWinResultCount + 1
                If lngMwtWinResultCount > lngMwtWinDimCount Then
                    lngMwtWinDimCount = lngMwtWinDimCount + DIM_CHUNK
                    ReDim Preserve strPeptideNameMwtWin(lngMwtWinDimCount)
                End If
                
            Next lngResidueEnd
        Next lngResidueStart
        lngStopTime = GetTickCount()
        lngMwtWinWorkTime = lngStopTime - lngStartTime
        Debug.Print ""
        Debug.Print "MwtWin time (" & lngMwtWinResultCount & " peptides) = " & lngMwtWinWorkTime & " msec"
        
''        lngIcr2lsResultCount = 0
''        Debug.Print "Starting residue is ";
''        lngStartTime = GetTickCount()
''        For lngResidueStart = 1 To Len(strProtein)
''            If lngResidueStart Mod 10 = 0 Then
''                Debug.Print lngResidueStart & ", ";
''                DoEvents
''            End If
''            ' Use DoEvents on every iteration since Icr2ls is quite slow
''            DoEvents
''
''            For lngResidueEnd = 1 To Len(strProtein) - lngResidueStart
''                If lngResidueEnd - lngResidueStart > 50 Then
''                    Exit For
''                End If
''
''                strPeptideResidues = Mid(strProtein, lngResidueStart, lngResidueEnd)
''                strPeptideNameIcr2ls(lngIcr2lsResultCount) = ICRTools.TrypticName(strProtein, strPeptideResidues)
''
''                lngIcr2lsResultCount = lngIcr2lsResultCount + 1
''                If lngIcr2lsResultCount > lngICR2lsDimCount Then
''                    lngICR2lsDimCount = lngICR2lsDimCount + DIM_CHUNK
''                    ReDim Preserve strPeptideNameIcr2ls(lngICR2lsDimCount)
''                End If
''            Next lngResidueEnd
''        Next lngResidueStart
''        lngStopTime = GetTickCount()
''        lngIcr2lsWorkTime = lngStopTime - lngStartTime
''        Debug.Print ""
''        Debug.Print "Icr2ls time (" & lngMwtWinResultCount & " peptides) = " & lngIcr2lsWorkTime & " msec"
            
''        ' Check that results match
''        For lngIndex = 0 To lngMwtWinResultCount - 1
''            If Left(strPeptideNameMwtWin(lngIndex), 1) = "t" Then
''                If Val(Right(strPeptideNameMwtWin(lngIndex), 1)) < 5 Then
''                    ' Icr2LS does not return the correct name when strPeptideResidues contains 5 or more tryptic peptides
''                    If strPeptideNameMwtWin(lngIndex) <> strPeptideNameIcr2ls(lngIndex) Then
''                        frmResults.AddToResults "Difference found, index = " & lngIndex & ", " & strPeptideNameMwtWin(lngIndex) & " vs. " & strPeptideNameIcr2ls(lngIndex)
''                        blnDifferenceFound = True
''                    End If
''                End If
''            Else
''                If strPeptideNameMwtWin(lngIndex) <> strPeptideNameIcr2ls(lngIndex) Then
''                    frmResults.AddToResults "Difference found, index = " & lngIndex & ", " & strPeptideNameMwtWin(lngIndex) & " vs. " & strPeptideNameIcr2ls(lngIndex)
''                    blnDifferenceFound = True
''                End If
''            End If
''        Next lngIndex
    
    Next lngMultipleIteration
    
    frmResults.AddToResults "Check of Tryptic Sequence functions Complete"
        
    Me.MousePointer = vbNormal
End Sub

Private Sub UpdateResultsForCompound(objCompound As MWCompoundClass)
    With objCompound
        If .ErrorDescription = "" Then
            txtFormula = .FormulaCapitalized
            FindMass
        Else
            lblStatus = .ErrorDescription
        End If
    End With
    
End Sub
Private Sub cboStdDevMode_Click()
    Select Case cboStdDevMode.ListIndex
    Case 1
        objMwtWin.StdDevMode = smScientific
    Case 2
        objMwtWin.StdDevMode = smDecimal
    Case Else
        objMwtWin.StdDevMode = smShort
    End Select

End Sub

Private Sub cboWeightMode_Click()
    Select Case cboWeightMode.ListIndex
    Case 1
        objMwtWin.SetElementMode emIsotopicMass
    Case 2
        objMwtWin.SetElementMode emIntegerMass
    Case Else
        objMwtWin.SetElementMode emAverageMass
    End Select
End Sub

Private Sub cmdClose_Click()
    Set objMwtWin = Nothing
    Unload Me
    End
End Sub

Private Sub cmdConvertToEmpirical_Click()
    With objMwtWin.Compound
        .Formula = txtFormula
        .ConvertToEmpirical
    End With
     
    UpdateResultsForCompound objMwtWin.Compound
 
End Sub

Private Sub cmdExpandAbbreviations_Click()
    With objMwtWin.Compound
        .Formula = txtFormula
        .ExpandAbbreviations
    End With

    UpdateResultsForCompound objMwtWin.Compound
End Sub

Private Sub cmdFindMass_Click()
    FindMass
    FindPercentComposition
End Sub

Private Sub cmdTestFunctions_Click()
    TestAccessFunctions
End Sub

Private Sub cmdTestGetTrypticName_Click()
    TestTrypticName
End Sub

Private Sub Form_Load()
    PopulateComboBoxes
End Sub

