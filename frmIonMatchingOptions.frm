VERSION 5.00
Begin VB.Form frmIonMatchOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ion Matching Options"
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8250
   HelpContextID   =   3085
   Icon            =   "frmIonMatchingOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   8250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Tag             =   "14000"
   Begin VB.Frame fraPlottingOptions 
      Caption         =   "Spectrum Plotting Options"
      Height          =   1935
      Left            =   3960
      TabIndex        =   26
      Tag             =   "14240"
      Top             =   4920
      Width           =   4215
      Begin VB.CommandButton cmdEditAutoLabelPeaksOptions 
         Caption         =   "Edit Auto-Label Options"
         Height          =   360
         Left            =   960
         TabIndex        =   37
         Tag             =   "14020"
         Top             =   1440
         Width           =   2115
      End
      Begin VB.CheckBox chkAutoLabelMass 
         Caption         =   "Auto-Label Peaks on Matching Ion Spectrum"
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Tag             =   "14270"
         Top             =   1080
         Width           =   3495
      End
      Begin VB.CheckBox chkPlotSpectrumInverted 
         Caption         =   "Plot Fragmentation Spectrum Inverted"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Tag             =   "14260"
         Top             =   840
         Width           =   3495
      End
      Begin VB.Label lblMatchingIonDataColorLabel 
         BackStyle       =   0  'Transparent
         Caption         =   "Matching Ion Data Color"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Tag             =   "14255"
         Top             =   525
         Width           =   2775
      End
      Begin VB.Label lblMatchingIonDataColor 
         BackColor       =   &H0000C000&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3120
         TabIndex        =   31
         Tag             =   "14257"
         ToolTipText     =   "Click to change"
         Top             =   525
         Width           =   375
      End
      Begin VB.Label lblFragSpectrumColorLabel 
         BackStyle       =   0  'Transparent
         Caption         =   "Fragmentation Data Color"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Tag             =   "14250"
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label lblFragSpectrumColor 
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3120
         TabIndex        =   29
         Tag             =   "14257"
         ToolTipText     =   "Click to change"
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame fraFragSpecLabels 
      Caption         =   "Frgamentation Spectrum Labels"
      Height          =   1095
      Left            =   120
      TabIndex        =   22
      Tag             =   "14200"
      Top             =   5280
      Width           =   3735
      Begin VB.CheckBox chkFragSpecEmphasizeProlineYIons 
         Caption         =   "Emphasize Proline y ions"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Tag             =   "14230"
         Top             =   720
         Width           =   3495
      End
      Begin VB.CheckBox chkFragSpecLabelOtherIons 
         Caption         =   "Label neutral loss ions"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Tag             =   "14220"
         Top             =   480
         Width           =   3495
      End
      Begin VB.CheckBox chkFragSpecLabelMainIons 
         Caption         =   "Label main ions (a, b, y)"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Tag             =   "14210"
         Top             =   240
         Value           =   1  'Checked
         Width           =   3495
      End
   End
   Begin VB.CommandButton cmdResetToDefaults 
      Caption         =   "&Reset to Defaults"
      Height          =   480
      Left            =   480
      TabIndex        =   34
      Tag             =   "14010"
      Top             =   6480
      Width           =   1275
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "Cl&ose"
      Height          =   480
      Left            =   2160
      TabIndex        =   35
      Tag             =   "4000"
      Top             =   6480
      Width           =   1155
   End
   Begin VB.Frame fraNormalization 
      Caption         =   "Normalization Options for Imported Data"
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Tag             =   "14100"
      Top             =   120
      Width           =   3735
      Begin VB.TextBox txtGroupIonMassWindow 
         Height          =   285
         Left            =   2280
         TabIndex        =   3
         Text            =   "0.5"
         Top             =   540
         Width           =   735
      End
      Begin VB.CheckBox chkGroupSimilarIons 
         Caption         =   "&Group Similar Ions (Bin Data)"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Tag             =   "14110"
         Top             =   240
         Width           =   3375
      End
      Begin VB.TextBox txtMassRegions 
         Height          =   285
         Left            =   2760
         TabIndex        =   8
         Text            =   "1"
         Top             =   1440
         Width           =   800
      End
      Begin VB.TextBox txtIonCountToUse 
         Height          =   285
         Left            =   2760
         TabIndex        =   6
         Text            =   "200"
         Top             =   1080
         Width           =   800
      End
      Begin VB.TextBox txtNormalizedIntensity 
         Height          =   285
         Left            =   2760
         TabIndex        =   10
         Text            =   "100"
         Top             =   1800
         Width           =   800
      End
      Begin VB.Label lblGroupSimilarIons 
         Caption         =   "Mass Window"
         Height          =   255
         Left            =   480
         TabIndex        =   2
         Tag             =   "14115"
         Top             =   570
         Width           =   1815
      End
      Begin VB.Label lblMassRegions 
         Caption         =   "Mass region subdivisions"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Tag             =   "14140"
         Top             =   1440
         Width           =   2655
      End
      Begin VB.Label lblIonCountToUse 
         Caption         =   "Number of Ions to Use"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Tag             =   "14130"
         Top             =   1080
         Width           =   2655
      End
      Begin VB.Label lblNormalizedIntensity 
         Caption         =   "Normalized Intensity"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Tag             =   "14120"
         Top             =   1800
         Width           =   2655
      End
      Begin VB.Label lblGroupSimilarIonsUnits 
         Caption         =   "Da"
         Height          =   255
         Left            =   3120
         TabIndex        =   4
         Tag             =   "12350"
         Top             =   600
         Width           =   495
      End
   End
   Begin VB.Frame fraIonIntensities 
      Caption         =   "Ion Intensities of Predicted Ions"
      Height          =   2655
      Left            =   120
      TabIndex        =   11
      Tag             =   "14150"
      Top             =   2520
      Width           =   3735
      Begin VB.TextBox txtIonIntensity 
         Height          =   285
         Index           =   4
         Left            =   2760
         TabIndex        =   39
         Text            =   "50"
         Top             =   1600
         Width           =   800
      End
      Begin VB.TextBox txtIonIntensity 
         Height          =   285
         Index           =   3
         Left            =   2760
         TabIndex        =   38
         Text            =   "50"
         Top             =   1275
         Width           =   800
      End
      Begin VB.TextBox txtIonIntensity 
         Height          =   285
         Index           =   0
         Left            =   2760
         TabIndex        =   13
         Text            =   "10"
         Top             =   300
         Width           =   800
      End
      Begin VB.TextBox txtNeutralLosses 
         Height          =   285
         Left            =   2760
         TabIndex        =   21
         Text            =   "10"
         Top             =   2250
         Width           =   800
      End
      Begin VB.TextBox txtBYIonShoulders 
         Height          =   285
         Left            =   2760
         TabIndex        =   19
         Text            =   "25"
         Top             =   1925
         Width           =   800
      End
      Begin VB.TextBox txtIonIntensity 
         Height          =   285
         Index           =   2
         Left            =   2760
         TabIndex        =   17
         Text            =   "50"
         Top             =   950
         Width           =   800
      End
      Begin VB.TextBox txtIonIntensity 
         Height          =   285
         Index           =   1
         Left            =   2760
         TabIndex        =   15
         Text            =   "50"
         Top             =   625
         Width           =   800
      End
      Begin VB.Label lblIonIntensity 
         Caption         =   "Z Ion Intensity"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   41
         Tag             =   "14170"
         Top             =   1600
         Width           =   2655
      End
      Begin VB.Label lblIonIntensity 
         Caption         =   "C Ion Intensity"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   40
         Tag             =   "14170"
         Top             =   1275
         Width           =   2655
      End
      Begin VB.Label lblIonIntensity 
         Caption         =   "A Ion Intensity"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Tag             =   "14160"
         Top             =   300
         Width           =   2660
      End
      Begin VB.Label lblNeutralLosses 
         Caption         =   "Neutral Losses"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Tag             =   "14190"
         Top             =   2250
         Width           =   2655
      End
      Begin VB.Label lblBYIonShoulders 
         Caption         =   "B/Y Ion shoulders"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Tag             =   "14180"
         Top             =   1925
         Width           =   2655
      End
      Begin VB.Label lblIonIntensity 
         Caption         =   "Y Ion Intensity"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   16
         Tag             =   "14170"
         Top             =   950
         Width           =   2660
      End
      Begin VB.Label lblIonIntensity 
         Caption         =   "B Ion Intensity"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Tag             =   "14165"
         Top             =   625
         Width           =   2660
      End
   End
   Begin VB.Label lblPredictedIntensityExplanation 
      Caption         =   $"frmIonMatchingOptions.frx":08CA
      Height          =   2055
      Left            =   4080
      TabIndex        =   33
      Top             =   2640
      Width           =   3975
   End
   Begin VB.Label lblNormalizationExplanation 
      Caption         =   $"frmIonMatchingOptions.frx":0A49
      Height          =   2175
      Left            =   4080
      TabIndex        =   32
      Top             =   240
      Width           =   3975
   End
End
Attribute VB_Name = "frmIonMatchOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub EnableDisableControls()
    Dim boolGroupIons As Boolean
    
On Error GoTo EnableDisableControlsErrorHandler

    boolGroupIons = cChkBox(chkGroupSimilarIons.value)
    lblGroupSimilarIons.Enabled = boolGroupIons
    txtGroupIonMassWindow.Enabled = boolGroupIons
    lblGroupSimilarIonsUnits.Enabled = boolGroupIons
    
    Exit Sub

EnableDisableControlsErrorHandler:
    GeneralErrorHandler "frmIonMatchOptions|EnableDisableControls", Err.Number, Err.Description
    
End Sub

Public Sub SetDefaultValues()
    txtIonIntensity(0) = "10"
    txtIonIntensity(1) = "50"
    txtIonIntensity(2) = "50"
    txtBYIonShoulders = "25"
    txtNeutralLosses = "10"

    With frmFragmentationModelling
        .cboDoubleCharge.ListIndex = 8
        .chkRemovePrecursorIon.value = vbChecked
        .txtPrecursorMassWindow = "2"
        .txtIonMatchingWindow = "0.5"
    End With
    chkGroupSimilarIons.value = vbUnchecked
    txtGroupIonMassWindow = "0.5"
    txtNormalizedIntensity = "100"
    txtIonCountToUse = "200"
    txtMassRegions = "1"
    
    chkFragSpecLabelMainIons.value = vbChecked
    chkFragSpecLabelOtherIons.value = vbUnchecked
    chkFragSpecEmphasizeProlineYIons.value = vbUnchecked
    
    chkPlotSpectrumInverted.value = vbChecked
    chkAutoLabelMass = vbUnchecked
    
    lblFragSpectrumColor.BackColor = RGB(0, 0, 255)
    lblMatchingIonDataColor.BackColor = RGB(0, 128, 0)
    
End Sub

Private Sub chkFragSpecEmphasizeProlineYIons_Click()
    frmFragmentationModelling.ResetPredictedSpectrumCRC
    frmFragmentationModelling.DisplaySpectra
    
    On Error Resume Next
    chkFragSpecEmphasizeProlineYIons.SetFocus
End Sub

Private Sub chkFragSpecLabelMainIons_Click()
    frmFragmentationModelling.ResetPredictedSpectrumCRC
    frmFragmentationModelling.DisplaySpectra
    
    On Error Resume Next
    chkFragSpecLabelMainIons.SetFocus
End Sub

Private Sub chkFragSpecLabelOtherIons_Click()
    frmFragmentationModelling.ResetPredictedSpectrumCRC
    frmFragmentationModelling.DisplaySpectra
    
    On Error Resume Next
    chkFragSpecLabelOtherIons.SetFocus
End Sub

Private Sub chkGroupSimilarIons_Click()
    EnableDisableControls
    frmFragmentationModelling.UpdateIonMatchListWrapper
End Sub

Private Sub chkPlotSpectrumInverted_Click()
    frmFragmentationModelling.ResetPredictedSpectrumCRC
    frmFragmentationModelling.SetNeedToZoomOutFull True
    frmFragmentationModelling.UpdateIonMatchListWrapper
   
    On Error Resume Next
    chkPlotSpectrumInverted.SetFocus
    
End Sub

Private Sub cmdEditAutoLabelPeaksOptions_Click()
    frmFragmentationModelling.ShowAutoLabelPeaksOptions
End Sub

Private Sub cmdOK_Click()
    Me.Hide
End Sub

Private Sub cmdResetToDefaults_Click()
    SetDefaultValues
End Sub

Private Sub Form_Activate()
    Dim strMessage As String
    
    strMessage = LookupLanguageCaption(14050, "When a list of ions is imported into the program, ions of similar mass may optionally be grouped together via a binning process to reduce the total number of data points.  Next, ions around the precursor ion may be removed.")
    strMessage = strMessage & LookupLanguageCaption(14055, "  Then, the intensities are normalized to the given maximum intensity and ordered by decreasing intensity.  The top-most ions (number of ions to use) are divided into distinct mass regions and the ions in each region again normalized.")
    lblNormalizationExplanation.Caption = strMessage
    
    strMessage = LookupLanguageCaption(14060, "The masses of the predicted ions for a given peptide sequence are easily computed.  However, intensity values must also be assigned to the masses.")
    strMessage = strMessage & LookupLanguageCaption(14065, "  The B and Y ions are typically assigned the same intensity while the A ion is typically 5 times less intense.  Shoulder ions (masses ± 1 Da from the B and Y ions) can be added, in addition to including neutral losses (H2O, NH3, and PO4).")
    lblPredictedIntensityExplanation.Caption = strMessage
End Sub

Private Sub Form_Load()
    
    On Error GoTo IonMatchOptionsErrorHandler

    SizeAndCenterWindow Me, cWindowTopCenter, 8350, 7635
    
    EnableDisableControls
    
    Exit Sub

IonMatchOptionsErrorHandler:
    GeneralErrorHandler "frmIonMatchOptions|Form_Load", Err.Number, Err.Description

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    QueryUnloadFormHandler Me, Cancel, UnloadMode
End Sub

Private Sub lblFragSpectrumColor_Click()
    SelectCustomColor Me, lblFragSpectrumColor
End Sub

Private Sub lblMatchingIonDataColor_Click()
    SelectCustomColor Me, lblMatchingIonDataColor
End Sub

Private Sub txtBYIonShoulders_GotFocus()
    HighlightOnFocus txtBYIonShoulders
End Sub

Private Sub txtBYIonShoulders_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtBYIonShoulders, KeyAscii, True
End Sub

Private Sub txtBYIonShoulders_Validate(Cancel As Boolean)
    ValidateTextboxValueDbl txtBYIonShoulders, 0, 32000, 25
End Sub

Private Sub txtGroupIonMassWindow_GotFocus()
    HighlightOnFocus txtGroupIonMassWindow
End Sub

Private Sub txtGroupIonMassWindow_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then frmFragmentationModelling.UpdateIonMatchListWrapper
End Sub

Private Sub txtGroupIonMassWindow_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtGroupIonMassWindow, KeyAscii, True, True, False
End Sub

Private Sub txtGroupIonMassWindow_Validate(Cancel As Boolean)
    ValidateTextboxValueDbl txtGroupIonMassWindow, 0, 100, 0.5
    If Val(GetMostRecentTextBoxValue) <> Val(txtGroupIonMassWindow) Then frmFragmentationModelling.UpdateIonMatchListWrapper
End Sub

Private Sub txtIonCountToUse_GotFocus()
    HighlightOnFocus txtIonCountToUse
End Sub

Private Sub txtIonCountToUse_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then frmFragmentationModelling.UpdateIonMatchListWrapper
End Sub

Private Sub txtIonCountToUse_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtIonCountToUse, KeyAscii, True
    frmFragmentationModelling.SetFragMatchSettingsChanged
End Sub

Private Sub txtIonCountToUse_Validate(Cancel As Boolean)
    ValidateTextboxValueDbl txtIonCountToUse, 10, 32000, 200
    If Val(GetMostRecentTextBoxValue) <> Val(txtIonCountToUse) Then frmFragmentationModelling.UpdateIonMatchListWrapper
End Sub

Private Sub txtIonIntensity_GotFocus(Index As Integer)
    HighlightOnFocus txtIonIntensity(Index)
End Sub

Private Sub txtIonIntensity_KeyPress(Index As Integer, KeyAscii As Integer)
    TextBoxKeyPressHandler txtIonIntensity(Index), KeyAscii, True
End Sub

Private Sub txtIonIntensity_Validate(Index As Integer, Cancel As Boolean)
    Dim dblDefaultIntensity As Double
    
    If Index = 0 Then
        dblDefaultIntensity = 10
    Else
        dblDefaultIntensity = 50
    End If
        
    ValidateTextboxValueDbl txtIonIntensity(Index), 0, 32000, dblDefaultIntensity
    
End Sub

Private Sub txtMassRegions_GotFocus()
    HighlightOnFocus txtMassRegions
End Sub

Private Sub txtMassRegions_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then frmFragmentationModelling.UpdateIonMatchListWrapper
End Sub

Private Sub txtMassRegions_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtMassRegions, KeyAscii, True
    frmFragmentationModelling.SetFragMatchSettingsChanged
End Sub

Private Sub txtMassRegions_Validate(Cancel As Boolean)
    ValidateTextboxValueDbl txtMassRegions, 0, 1000, 1
    If Val(GetMostRecentTextBoxValue) <> Val(txtMassRegions) Then frmFragmentationModelling.UpdateIonMatchListWrapper
End Sub

Private Sub txtNeutralLosses_GotFocus()
    HighlightOnFocus txtNeutralLosses
End Sub

Private Sub txtNeutralLosses_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtNeutralLosses, KeyAscii, True
End Sub

Private Sub txtNeutralLosses_Validate(Cancel As Boolean)
    ValidateTextboxValueDbl txtNeutralLosses, 0, 32000, 10
End Sub

Private Sub txtNormalizedIntensity_GotFocus()
    HighlightOnFocus txtNormalizedIntensity
End Sub

Private Sub txtNormalizedIntensity_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then frmFragmentationModelling.UpdateIonMatchListWrapper
End Sub

Private Sub txtNormalizedIntensity_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtNormalizedIntensity, KeyAscii, True
    frmFragmentationModelling.SetFragMatchSettingsChanged
End Sub

Private Sub txtNormalizedIntensity_Validate(Cancel As Boolean)
    ValidateTextboxValueDbl txtNormalizedIntensity, 0, 32000, 100
    If Val(GetMostRecentTextBoxValue) <> Val(txtNormalizedIntensity) Then frmFragmentationModelling.UpdateIonMatchListWrapper
End Sub
