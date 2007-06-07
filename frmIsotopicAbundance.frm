VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmIsotopicDistribution 
   Caption         =   "Isotopic Distribution"
   ClientHeight    =   5205
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10830
   HelpContextID   =   3100
   Icon            =   "frmIsotopicAbundance.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5205
   ScaleWidth      =   10830
   StartUpPosition =   3  'Windows Default
   Tag             =   "15000"
   Begin VB.TextBox txtChargeState 
      Height          =   285
      Left            =   4920
      TabIndex        =   7
      Text            =   "1"
      Top             =   915
      Width           =   615
   End
   Begin VB.Frame fraIonComparisonList 
      Caption         =   "Comparison List"
      Height          =   2175
      Left            =   5640
      TabIndex        =   22
      Tag             =   "15195"
      Top             =   2880
      Width           =   4575
      Begin VB.CheckBox chkComparisonListNormalize 
         Caption         =   "Normalize pasted ion list"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Tag             =   "15105"
         Top             =   1000
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin VB.CommandButton cmdComparisonListClear 
         Caption         =   "Clear list"
         Height          =   375
         Left            =   2880
         TabIndex        =   28
         Tag             =   "15165"
         Top             =   1320
         Width           =   1455
      End
      Begin VB.CommandButton cmdComparisonListPaste 
         Caption         =   "Paste list of ions to plot"
         Height          =   375
         Left            =   120
         TabIndex        =   27
         Tag             =   "15160"
         Top             =   1320
         Width           =   2535
      End
      Begin VB.ComboBox cboComparisonListPlotType 
         Height          =   315
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Tag             =   "15140"
         Top             =   630
         Width           =   2295
      End
      Begin VB.Label lblComparisonListDataPoints 
         Caption         =   "0"
         Height          =   255
         Left            =   2640
         TabIndex        =   30
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label lblComparisonListDataPointsLabel 
         Caption         =   "Comparison list data points:"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Tag             =   "15170"
         Top             =   1800
         Width           =   2415
      End
      Begin VB.Label lblComparisonListPlotColorLabel 
         BackStyle       =   0  'Transparent
         Caption         =   "Plot Data Color"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Tag             =   "15120"
         Top             =   300
         Width           =   1935
      End
      Begin VB.Label lblComparisonListPlotColor 
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2040
         TabIndex        =   24
         Tag             =   "15125"
         ToolTipText     =   "Click to change"
         Top             =   300
         Width           =   375
      End
      Begin VB.Label lblComparisonListPlotType 
         BackStyle       =   0  'Transparent
         Caption         =   "Plot Type"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Tag             =   "15130"
         Top             =   660
         Width           =   1935
      End
   End
   Begin VB.Frame fraPlotOptions 
      Caption         =   "Options"
      Height          =   2655
      Left            =   5640
      TabIndex        =   9
      Tag             =   "15190"
      Top             =   120
      Width           =   4575
      Begin VB.TextBox txtEffectiveResolutionMass 
         Height          =   300
         Left            =   3000
         TabIndex        =   19
         Text            =   "1000"
         Top             =   1770
         Width           =   855
      End
      Begin VB.TextBox txtEffectiveResolution 
         Height          =   300
         Left            =   3000
         TabIndex        =   17
         Text            =   "5000"
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox txtGaussianQualityFactor 
         Height          =   300
         Left            =   3000
         TabIndex        =   21
         Text            =   "50"
         Top             =   2190
         Width           =   855
      End
      Begin VB.ComboBox cboPlotType 
         Height          =   315
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Tag             =   "15140"
         Top             =   960
         Width           =   2295
      End
      Begin VB.CheckBox chkPlotResults 
         Caption         =   "&Plot Results"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Tag             =   "15100"
         Top             =   300
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CheckBox chkAutoLabelPeaks 
         Caption         =   "&Auto-Label Peaks"
         Height          =   255
         Left            =   2040
         TabIndex        =   11
         Tag             =   "15105"
         Top             =   300
         Width           =   2295
      End
      Begin VB.Label lblEffectiveResolutionMass 
         BackStyle       =   0  'Transparent
         Caption         =   "Effective Resolution M/Z"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Tag             =   "15130"
         Top             =   1800
         Width           =   2655
      End
      Begin VB.Label lblEffectiveResolution 
         BackStyle       =   0  'Transparent
         Caption         =   "Effective Resolution"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Tag             =   "15130"
         Top             =   1470
         Width           =   2655
      End
      Begin VB.Label lblGaussianQualityFactor 
         BackStyle       =   0  'Transparent
         Caption         =   "Gaussian Quality Factor"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Tag             =   "15130"
         Top             =   2220
         Width           =   2655
      End
      Begin VB.Label lblPlotType 
         BackStyle       =   0  'Transparent
         Caption         =   "Plot Type"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Tag             =   "15130"
         Top             =   990
         Width           =   1935
      End
      Begin VB.Label lblPlotColor 
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2040
         TabIndex        =   13
         Tag             =   "15125"
         ToolTipText     =   "Click to change"
         Top             =   630
         Width           =   375
      End
      Begin VB.Label lblPlotColorLabel 
         BackStyle       =   0  'Transparent
         Caption         =   "Plot Data Color"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Tag             =   "15120"
         Top             =   630
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Cop&y"
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Tag             =   "10340"
      Top             =   840
      Width           =   1095
   End
   Begin RichTextLib.RichTextBox rtfFormula 
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Tag             =   "5050"
      ToolTipText     =   "Type the molecular formula here."
      Top             =   120
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   873
      _Version        =   393217
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      TextRTF         =   $"frmIsotopicAbundance.frx":08CA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtResults 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Tag             =   "15150"
      ToolTipText     =   "Isotopic distribution results"
      Top             =   1680
      Width           =   5415
   End
   Begin VB.CommandButton cmdCompute 
      Caption         =   "&Compute"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Tag             =   "15110"
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Cl&ose"
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Tag             =   "4000"
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label lblChargeState 
      Alignment       =   1  'Right Justify
      Caption         =   "C&harge state:"
      Height          =   255
      Left            =   3480
      TabIndex        =   6
      Tag             =   "15070"
      Top             =   930
      Width           =   1335
   End
   Begin VB.Label lblResults 
      Caption         =   "Results:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Tag             =   "15060"
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label lblFormula 
      Caption         =   "&Formula:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Tag             =   "15050"
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "frmIsotopicDistribution"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents objSpectrum As CWSpectrumDLL.Spectrum
Attribute objSpectrum.VB_VarHelpID = -1

Private Const MAX_SERIES_COUNT = 2
Private Const GAUSSIAN_CONVERSION_DATACOUNT_WARNING_THRESHOLD = 1000

Private Enum ipmIsotopicPlotModeConstants
    ipmSticksToZero = 0
    ipmGaussian = 1
    ipmContinuousData = 2                           ' Used for raw, uncentered data
End Enum

Private Enum lcwfLastControlWithFocusConstants
    lcwfFormula = 0
    lcwfEffectiveResolution = 1
    lcwfEffectiveResolutionMass = 2
    lcwfGaussianQualityFactor = 3
    lcwfChargeState = 4
End Enum

Private ConvolutedMSDataCount As Long
Private ConvolutedMSData2D() As Double              ' 1-based array (since it is 1-based in the Dll)

Private ComparisonListDataCount As Long
Private ComparisonListXVals() As Double          ' 0-based array
Private ComparisonListYVals() As Double          ' 0-based array

Private CRCForPredictedSpectrumSaved As Double
Private CRCForComparisonIonListSaved As Double
Private mNeedToZoomOutFull As Boolean

Private mLastControlWithFocus As lcwfLastControlWithFocusConstants

Private mFormActivatedByUser As Boolean      ' Set true the first time the user activates the form; always true from then on
                                             ' Used to prevent plotting the mass spectrum until after the user has activated the isotopic distribution form at least once

Private mUserWarnedLargeDataSetForGaussianConversion As Boolean
Private mDelayUpdatingPlot As Boolean

Private Sub ComparisonIonListClear()
    ComparisonListDataCount = 0
    ReDim ComparisonListXVals(0)
    ReDim ComparisonListYVals(0)
    UpdateComparisonListStatus
    
    PlotIsotopicDistribution
    
End Sub

Private Sub ComparisonIonListPaste()
    
    Dim strIonList As String, strLineWork As String
    Dim blnAllowCommaDelimeter As Boolean
    Dim lngCrLfLoc As Long, lngValuesToPopulate As Long
    Dim lngIndex As Long, lngStartIndex As Long
    Dim strDelimeter As String, intDelimeterLength As Integer
    
    Dim strLeftChar As String, strDelimeterList As String
    Dim dblParsedVals() As Double      ' 0-based array
    Dim intParseValCount As Integer
    Dim dblMaximumIntensity As Double
    
On Error GoTo ComparisonIonListPasteErrorHandler

''    ' Warn user about replacing data
''    If Not QueryExistingIonsInList() Then Exit Sub
    
    ' Grab text from clipboard
    strIonList = GetClipboardTextSmart()
    
    If Len(strIonList) = 0 Then
        MsgBox LookupMessage(980), vbInformation + vbOKOnly, LookupMessage(985)
        Exit Sub
    End If
        
    ' Determine if commas are used for the decimal point in this locale
    If DetermineDecimalPoint() = "," Then
        blnAllowCommaDelimeter = False
    Else
        blnAllowCommaDelimeter = True
    End If
        
    frmProgress.InitializeForm LookupMessage(990), 0, Len(strIonList)
    frmProgress.ToggleAlwaysOnTop True
    frmProgress.UpdateCurrentSubTask LookupMessage(1000)
    
    ' First Determine number of data points in strIonList
    ' This isn't necessary (while it is necessary in frmFragmentationModelling.PasteIonMatchList),
    '  but it doesn't take that much time and allows this code to be nearly identical to the frmFragmentationModelling code
    
    ' First look for the first occurrence of a valid delimeter
    lngCrLfLoc = ParseStringFindCrlfIndex(strIonList, intDelimeterLength)
    
    If lngCrLfLoc > 0 Then
        ' Record the first character of the delimeter (if vbCrLf, then recording vbCr), otherwise, recording just vbCr or just vbLf)
        strDelimeter = Mid(strIonList, lngCrLfLoc, 1)
        
        ' Now determine the number of occurrences of the delimeter
        lngValuesToPopulate = 0
        For lngIndex = 1 To Len(strIonList)
            If lngIndex Mod 500 = 0 Then
                frmProgress.UpdateProgressBar lngIndex
                If KeyPressAbortProcess > 1 Then Exit For
            End If
            
            If Mid(strIonList, lngIndex, 1) = strDelimeter Then
                lngValuesToPopulate = lngValuesToPopulate + 1
            End If
        Next lngIndex
        
        If lngValuesToPopulate > 1 And KeyPressAbortProcess = 0 Then
            ' Update frmProgress to use the correct progress bar size
            frmProgress.InitializeForm LookupMessage(990), 0, lngValuesToPopulate
            frmProgress.UpdateCurrentSubTask LookupMessage(1010)
        End If
        
        ' Initialize IonMatchList
        ComparisonListDataCount = 0
        ReDim ComparisonListXVals(lngValuesToPopulate)
        ReDim ComparisonListYVals(lngValuesToPopulate)
                            
        ' Construct Delimeter List: Contains a space, Tab, and possibly comma
        strDelimeterList = " " & vbTab & vbCr & vbLf
        If blnAllowCommaDelimeter Then
            strDelimeterList = strDelimeterList & ","
        End If
        
        ' Initialize dblMaximumIntensity
        dblMaximumIntensity = LowestValueForDoubleDataType
        
        ' Actually parse the data
        ' I process the list using the Mid() function since this is the fastest method
        ' However, Mid() becomes slower when the index value reaches 1000 or more (roughly)
        '  so I discard already-parsed data from strIonList every 1000 characters (approximately)
        Do While Len(strIonList) > 0
            lngStartIndex = 1
            For lngIndex = 1 To Len(strIonList)
                If lngIndex Mod 100 = 0 Then
                    frmProgress.UpdateProgressBar ComparisonListDataCount
                    If KeyPressAbortProcess > 1 Then Exit For
                End If
                
                If Mid(strIonList, lngIndex, 1) = strDelimeter Then
                    strLineWork = Mid(strIonList, lngStartIndex, lngIndex - 1)
        
                    If Len(strLineWork) > 0 Then
                        
                        ' Only parse if the first letter (trimmed) is a number (may start with - or +)
                        strLeftChar = Left(strLineWork, 1)
                        If IsNumeric(strLeftChar) Or strLeftChar = "-" Or strLeftChar = "+" Then
                            intParseValCount = ParseStringValuesDbl(strLineWork, dblParsedVals(), 2, strDelimeterList, , False, False, False)
                        
                            If intParseValCount >= 2 Then
                                If dblParsedVals(0) <> 0 And dblParsedVals(0) <> 0 Then
                                    If dblParsedVals(1) > dblMaximumIntensity Then
                                        dblMaximumIntensity = dblParsedVals(1)
                                    End If
                                    
                                    ComparisonListXVals(ComparisonListDataCount) = dblParsedVals(0)
                                    ComparisonListYVals(ComparisonListDataCount) = dblParsedVals(1)
                                    ComparisonListDataCount = ComparisonListDataCount + 1
                                    
                                    
                                    If ComparisonListDataCount > lngValuesToPopulate Then
                                        ' This shouldn't happen
                                        Debug.Assert False
                                        Exit Do
                                    End If
                                End If
                            End If
                        End If
                    
                    End If
                    
                    lngStartIndex = lngIndex + intDelimeterLength
                    lngIndex = lngIndex + intDelimeterLength - 1
                    
                    If lngIndex > 1000 Then
                        ' Reduce the size of strIonList since the Mid() function gets slower with longer strings
                        strIonList = Mid(strIonList, lngIndex + 1)
                        lngIndex = 1
                        Exit For
                    End If
                End If
            Next lngIndex
            If lngIndex > Len(strIonList) Then Exit Do
        Loop
    End If
    
    frmProgress.HideForm
    If ComparisonListDataCount > 0 And KeyPressAbortProcess = 0 Then
        
        If dblMaximumIntensity <> 0 And cChkBox(chkComparisonListNormalize) Then
            ' Normalize the Comparison ion list to 100
            For lngIndex = 0 To ComparisonListDataCount - 1
                ComparisonListYVals(lngIndex) = ComparisonListYVals(lngIndex) / dblMaximumIntensity * 100
            Next lngIndex
        End If
        
        PlotIsotopicDistribution False
    Else
        If KeyPressAbortProcess = 0 Then
            MsgBox LookupMessage(1020), vbInformation + vbOKOnly, LookupMessage(985)
        Else
            MsgBox LookupMessage(940), vbInformation + vbOKOnly, LookupMessage(945)
        End If
    End If

    UpdateComparisonListStatus
    
    Exit Sub
    
ComparisonIonListPasteErrorHandler:
    frmProgress.HideForm
    UpdateComparisonListStatus
    GeneralErrorHandler "frmIsotopicDistribution|ComparisonIonListPaste", Err.Number, Err.Description

End Sub

Private Sub CopyResults()
        
    Clipboard.Clear
    Clipboard.SetText txtResults, vbCFText

End Sub

Public Sub EnableDisablePlotUpdates(blnDisableUpdating As Boolean)
    mDelayUpdatingPlot = blnDisableUpdating
    If Not blnDisableUpdating Then
        PlotIsotopicDistribution True
    End If
End Sub

Private Sub PlotIsotopicDistribution(Optional blnIgnoreCRC As Boolean = False)
    Dim lngIndex As Long
    
    Static blnSeriesPlotModeInitialized(MAX_SERIES_COUNT) As Boolean
    Static blnUpdatingPlot As Boolean
    
    Dim PredictedSpectrumCRC As Double
    Dim dblMassTimesIntensity As Double
    Dim blnUpdatePlot As Boolean
    Dim blnAllowResumeNextErrorHandling As Boolean
    
    Dim dblXVals() As Double            ' 0-based array
    Dim dblYVals() As Double            ' 0-based array
    Dim lngDataCount As Long
        
    Dim intSeriesNumber As Integer
    Dim strLegendCaption As String
    Dim eIsoPlotMode As ipmIsotopicPlotModeConstants
    Dim lngPlotColor As Long
    
    Dim eResponse As VbMsgBoxResult
    Dim strMessage As String
    Dim blnCursorVisibilitySaved As Boolean
    Dim blnAutoHideCaptionsSaved As Boolean
    
    If Not mFormActivatedByUser Or mDelayUpdatingPlot Then Exit Sub
    
On Error GoTo PlotIsotopicDistributionErrorHandler

    blnAllowResumeNextErrorHandling = True
    
    ' Compute a CRC value for the current PredictedSpectrum() array and compare to the previously computed value
    PredictedSpectrumCRC = 0
    For lngIndex = 1 To ConvolutedMSDataCount
        dblMassTimesIntensity = Abs((lngIndex) * ConvolutedMSData2D(lngIndex, 0) * ConvolutedMSData2D(lngIndex, 1))
        If dblMassTimesIntensity > 0 Then
            PredictedSpectrumCRC = PredictedSpectrumCRC + Log(dblMassTimesIntensity)
        End If
    Next lngIndex
    
    ' If the new CRC is different than the old one then re-plot the spectrun
    If PredictedSpectrumCRC <> CRCForPredictedSpectrumSaved Then
        CRCForPredictedSpectrumSaved = PredictedSpectrumCRC
        blnUpdatePlot = True
    End If
    
    ' Also compute a CRC value for the Comparison Ion List and compare to the previously computed value
    PredictedSpectrumCRC = 0
    For lngIndex = 0 To ComparisonListDataCount - 1
        dblMassTimesIntensity = Abs(lngIndex * ComparisonListXVals(lngIndex) * ComparisonListYVals(lngIndex))
        If dblMassTimesIntensity > 0 Then
            PredictedSpectrumCRC = PredictedSpectrumCRC + Log(dblMassTimesIntensity)
        End If
    Next lngIndex
    
    ' If the new CRC is different than the old one then re-plot the spectrun
    If PredictedSpectrumCRC <> CRCForComparisonIonListSaved Then
        CRCForComparisonIonListSaved = PredictedSpectrumCRC
        blnUpdatePlot = True
    End If
    
    If blnIgnoreCRC Then blnUpdatePlot = True
    
    If Not blnUpdatePlot Then
        Exit Sub
    End If
    
    If blnUpdatingPlot Then Exit Sub
    blnUpdatingPlot = True
    
    blnAllowResumeNextErrorHandling = False
    
    objSpectrum.ShowSpectrum
    
    ' Hide the Cursor and disable auto-hiding of annotations
    blnCursorVisibilitySaved = objSpectrum.GetCursorVisibility()
    If blnCursorVisibilitySaved Then objSpectrum.SetCursorVisible False

    blnAutoHideCaptionsSaved = objSpectrum.GetAnnotationDensityAutoHideCaptions()
    If blnAutoHideCaptionsSaved Then objSpectrum.SetAnnotationDensityAutoHideCaptions False, False
    
    If ConvolutedMSDataCount > 0 Then
        lngDataCount = ConvolutedMSDataCount
        ReDim dblXVals(0 To lngDataCount - 1)
        ReDim dblYVals(0 To lngDataCount - 1)
        
        ' Now fill dblXVals() and dblYVals()
        ' ConvolutedMSData2D() is 1-based, but dblXVals() and dblYVals() are 0-based
        For lngIndex = 1 To lngDataCount
            dblXVals(lngIndex - 1) = ConvolutedMSData2D(lngIndex, 0)
            dblYVals(lngIndex - 1) = ConvolutedMSData2D(lngIndex, 1)
        Next lngIndex
    Else
        lngDataCount = 0
    End If
    
    ' First plot the theoretical isotopic distribution
    intSeriesNumber = 1
    strLegendCaption = rtfFormula.Text
    eIsoPlotMode = cboPlotType.ListIndex
    lngPlotColor = lblPlotColor.BackColor
    PlotIsotopicDistributionWork objSpectrum, intSeriesNumber, lngDataCount, dblXVals(), dblYVals(), strLegendCaption, blnSeriesPlotModeInitialized(), eIsoPlotMode, lngPlotColor
    
    ' Now plot the comparison ion list (will clear series 2 if list is empty)
    intSeriesNumber = 2
    strLegendCaption = "Comparison List"
    eIsoPlotMode = cboComparisonListPlotType.ListIndex
    lngPlotColor = lblComparisonListPlotColor.BackColor
    
    ' Need to copy the data from CaprisonListXVals() to dblXVals() since the Gaussian conversion operation changes the array
    If ComparisonListDataCount > 0 Then
        lngDataCount = ComparisonListDataCount
        ReDim dblXVals(0 To UBound(ComparisonListXVals()))
        dblXVals() = ComparisonListXVals()
        dblYVals() = ComparisonListYVals()
    Else
        lngDataCount = 0
    End If
    
    ' Check for large number of data points when Gaussian mode is enabled
    If eIsoPlotMode = ipmGaussian And lngDataCount > GAUSSIAN_CONVERSION_DATACOUNT_WARNING_THRESHOLD Then
        If Not mUserWarnedLargeDataSetForGaussianConversion Then
            strMessage = LookupLanguageCaption(15300, "Warning, the ion comparison list contains a large number of data points.  Do you really want to convert each point to a Gaussian curve?")
            eResponse = MsgBox(strMessage, vbQuestion + vbYesNoCancel + vbDefaultButton2, LookupLanguageCaption(15305, "Large Data Point Count"))
            
            If eResponse = vbYes Then
                mUserWarnedLargeDataSetForGaussianConversion = True
            ElseIf eResponse = vbCancel Then
                lngDataCount = 0
            Else
                ' Switch to Continuous Data mode
                eIsoPlotMode = ipmContinuousData
                cboComparisonListPlotType.ListIndex = eIsoPlotMode
            End If
        End If
    End If
    
    PlotIsotopicDistributionWork objSpectrum, intSeriesNumber, lngDataCount, dblXVals(), dblYVals(), strLegendCaption, blnSeriesPlotModeInitialized(), eIsoPlotMode, lngPlotColor
    
    With objSpectrum
        If mNeedToZoomOutFull Then
            .ZoomOutFull
            mNeedToZoomOutFull = False
        End If
    
        .SetCursorPosition 100, 0
    
        ' Auto-hiding of captions was disabled above; re-enable if necessary
        .SetAnnotationDensityAutoHideCaptions blnAutoHideCaptionsSaved

        ' The cursor was hidden above; re-show it if necessary
        If blnCursorVisibilitySaved Then
            .SetCursorVisible True
        End If
        
        .SetSpectrumFormCurrentSeriesNumber 1
        
        .SetCustomTicksXAxis 0, 0, True
        .SetCustomTicksYAxis 0, 0, True
        
        .ShowSpectrum
    End With
    
    ' Return the focus back to this form (if possible)
    On Error Resume Next
    Select Case mLastControlWithFocus
    Case lcwfFormula
        rtfFormula.SetFocus
    Case lcwfEffectiveResolution
        txtEffectiveResolution.SetFocus
    Case lcwfEffectiveResolutionMass
        txtEffectiveResolutionMass.SetFocus
    Case lcwfGaussianQualityFactor
        txtGaussianQualityFactor.SetFocus
    Case lcwfChargeState
        txtChargeState.SetFocus
    Case Else
        Me.SetFocus
    End Select
    
    blnUpdatingPlot = False
    
    Exit Sub

PlotIsotopicDistributionErrorHandler:
    If blnAllowResumeNextErrorHandling Then
        ' Something is probably wrong with the CRC computation routines
        Debug.Assert False
        Resume Next
    Else
        GeneralErrorHandler "frmIsotopicDistribution|PlotIsotopicDistribution", Err.Number, Err.Description
    End If
    blnUpdatingPlot = False
    
End Sub

Private Sub PlotIsotopicDistributionWork(objThisCWSpectrum As CWSpectrumDLL.Spectrum, intSeriesNumber As Integer, lngDataCount As Long, dblXVals() As Double, dblYVals() As Double, strLegendCaption As String, blnSeriesPlotModeInitialized() As Boolean, eIsoPlotMode As ipmIsotopicPlotModeConstants, lngPlotColor As Long)

    Dim lngResolution As Long, lngResolutionMass As Long
    Dim lngQualityFactor As Long

    Dim udtAutoLabelPeaksSettings As udtAutoLabelPeaksOptionsType
    Dim dblMinimum As Double, dblMaximum As Double
    Dim ePlotMode As pmPlotModeConstants
    
On Error GoTo PlotIsotopicDistributionWorkErrorHandler

    If eIsoPlotMode = ipmGaussian Then
        lngResolution = CLngSafe(txtEffectiveResolution)
        lngResolutionMass = CLngSafe(txtEffectiveResolutionMass)
        lngQualityFactor = CIntSafe(txtGaussianQualityFactor)
        
        ValidateValueLng lngResolution, 1, 1000000000#, 5000
        ValidateValueLng lngQualityFactor, 1, 75, 50
        ValidateValueLng lngResolutionMass, 1, 1000000000#, 1000
        
        ' Note that dblXVals() and dblYVals() will be replaced with Gaussian representations of the peaks by the following function
        ConvertStickDataToGaussian2DArray Me, dblXVals(), dblYVals(), lngDataCount, lngResolution, lngResolutionMass, CInt(lngQualityFactor)
    End If
    
    If eIsoPlotMode = ipmSticksToZero Then
        ePlotMode = pmStickToZero
    Else
        ePlotMode = pmLines
    End If
    
    With objThisCWSpectrum
        If .GetSeriesCount() < intSeriesNumber Then
            .SetSeriesCount intSeriesNumber
        End If
        
        If Not blnSeriesPlotModeInitialized(intSeriesNumber) Then
            .SetSeriesPlotMode intSeriesNumber, pmStickToZero, False
            blnSeriesPlotModeInitialized(intSeriesNumber) = True
        End If
        
        .ClearData intSeriesNumber
        If lngDataCount > 0 Then
            ' Normal data
            .SetSeriesColor intSeriesNumber, lngPlotColor
            
            .SetDataXvsY intSeriesNumber, dblXVals(), dblYVals(), lngDataCount, strLegendCaption
        
            .SetSeriesPlotMode intSeriesNumber, ePlotMode, False
            
            If cChkBox(chkAutoLabelPeaks) Then
                udtAutoLabelPeaksSettings = .GetAutoLabelPeaksOptions()
                
                udtAutoLabelPeaksSettings.DisplayXPos = True
                udtAutoLabelPeaksSettings.IsContinuousData = (ePlotMode = pmLines)
                
                .SetAutoLabelPeaksOptions udtAutoLabelPeaksSettings
                .AutoLabelPeaks intSeriesNumber, udtAutoLabelPeaksSettings.DisplayXPos, udtAutoLabelPeaksSettings.DisplayYPos, udtAutoLabelPeaksSettings.CaptionAngle, udtAutoLabelPeaksSettings.IncludeArrow, udtAutoLabelPeaksSettings.HideInDenseRegions, udtAutoLabelPeaksSettings.PeakLabelCountMaximum
            End If
        
            .SetSpectrumFormCurrentSeriesNumber intSeriesNumber
            
            .GetRangeX dblMinimum, dblMaximum
            
            If dblXVals(0) < dblMinimum Or dblXVals(0) > dblMaximum Then mNeedToZoomOutFull = True
            If dblXVals(lngDataCount - 1) < dblMinimum Or dblXVals(lngDataCount - 1) > dblMaximum Then mNeedToZoomOutFull = True
        End If
    End With
    
    Exit Sub

PlotIsotopicDistributionWorkErrorHandler:
    GeneralErrorHandler "frmIsotopicDistribution|PlotIsotopicDistributionWork", Err.Number, Err.Description

End Sub

Private Sub PositionFormControls()
    Dim lngPreferredValue As Long
    
On Error GoTo PositionFormControlsErrorHandler

    With fraPlotOptions
        lngPreferredValue = Me.ScaleWidth - .Width - 120
        If lngPreferredValue < 5500 Then lngPreferredValue = 5500
        .Left = lngPreferredValue
    End With
    
    fraIonComparisonList.Left = fraPlotOptions.Left
    
    rtfFormula.Width = fraPlotOptions.Left - rtfFormula.Left - 120
    
    With txtResults
        .Top = lblResults.Top + 360
        .Left = 120
        
        lngPreferredValue = fraPlotOptions.Left - .Left - 120
        If lngPreferredValue < 2000 Then
            ' This shouldn't happen
            Debug.Assert False
            lngPreferredValue = 2000
        End If
        .Width = lngPreferredValue
        
        lngPreferredValue = Me.ScaleHeight - .Top - 120
        If lngPreferredValue < 1000 Then lngPreferredValue = 1000
        .Height = lngPreferredValue
    End With

    lngPreferredValue = txtResults.Left + txtResults.Width - txtChargeState.Width
    If lngPreferredValue < lblResults.Left + lblResults.Width + lblChargeState.Width Then
        lngPreferredValue = lblResults.Left + lblResults.Width + lblChargeState.Width
    End If
    txtChargeState.Left = lngPreferredValue
    lblChargeState.Left = txtChargeState.Left - lblChargeState.Width - 60
    
    Exit Sub

PositionFormControlsErrorHandler:
    Debug.Assert False
    Resume Next
    
End Sub

Private Function QueryExistingIonsInList() As Boolean
    Dim strMessage As String, eResponse As VbMsgBoxResult
    
    If ComparisonListDataCount > 0 Then
        strMessage = LookupMessage(910)
        eResponse = MsgBox(strMessage, vbQuestion + vbYesNoCancel + vbDefaultButton3, LookupMessage(920))
        If eResponse = vbYes Then
            QueryExistingIonsInList = True
        Else
            QueryExistingIonsInList = False
        End If
    Else
        QueryExistingIonsInList = True
    End If
    
End Function

Public Sub ResetCWSpectrumOptions()
    If cboComparisonListPlotType.ListIndex = ipmGaussian Then
        objSpectrum.ResetOptionsToDefaults True, True, 2, pmLines
    Else
        objSpectrum.ResetOptionsToDefaults True, True, 2, pmStickToZero
    End If
End Sub

Private Sub ResetPredictedSpectrumCRC()
    CRCForPredictedSpectrumSaved = 0
End Sub

Public Sub SetFormActivated()
    mFormActivatedByUser = True
End Sub

Public Sub SetNeedToZoomOutFull(blnEnable As Boolean)
    mNeedToZoomOutFull = blnEnable
End Sub

Public Sub StartIsotopicDistributionCalcs(blnSetFocusOnTextbox As Boolean, Optional blnSetFocusOnChargeTextbox As Boolean = False)
    Dim strResults As String, lngErrorID As Long
    Dim strFormula As String
    Dim strHeaderIsotopicAbundace As String, strHeaderMass As String
    Dim strHeaderFraction As String, strHeaderIntensity As String
    Dim lngCursorPos As Long
    
    Dim intChargeState As Integer
    
On Error GoTo StartIsotopicDistributionCalcsErrorHandler

    strFormula = rtfFormula.Text
    lngCursorPos = rtfFormula.SelStart
    
    If Len(Trim(strFormula)) = 0 Then Exit Sub
    
    RemoveHeightAdjustChar strFormula
    
    Me.MousePointer = vbHourglass
    
    If IsNumeric(txtChargeState) Then
        intChargeState = Val(txtChargeState)
        If intChargeState < 0 Then intChargeState = 1
        If intChargeState >= 10000 Then intChargeState = 10000
    Else
        intChargeState = 1
    End If
    
    strHeaderIsotopicAbundace = LookupLanguageCaption(15200, "Isotopic Abundances for")
    If intChargeState = 0 Then
        strHeaderMass = LookupLanguageCaption(15215, "Neutral Mass")
    Else
        strHeaderMass = LookupLanguageCaption(15210, "Mass/Charge")
    End If
    
    strHeaderFraction = LookupLanguageCaption(15220, "Fraction")
    strHeaderIntensity = LookupLanguageCaption(15230, "Intensity")
    
    ' Make sure we're using isotopic masses
    SwitchWeightMode emIsotopicMass
    
    ' Note: strFormula is passed ByRef
    lngErrorID = objMwtWin.ComputeIsotopicAbundances(strFormula, intChargeState, strResults, ConvolutedMSData2D(), ConvolutedMSDataCount, strHeaderIsotopicAbundace, strHeaderMass, strHeaderFraction, strHeaderIntensity)
    
    If lngErrorID = 0 Then
        ' Update rtfFormula with the capitalized formula
        rtfFormula = strFormula
        
        If cChkBox(chkPlotResults.value) Then
            PlotIsotopicDistribution
        End If
    ElseIf lngErrorID <> -1 Then
        MsgBox LookupMessage(lngErrorID), vbOKOnly + vbExclamation, LookupMessage(350)
    End If
    
    txtResults = strResults
    txtResults.Font = "Courier"
    
    rtfFormula = strFormula
    If lngCursorPos <= Len(strFormula) Then
        rtfFormula.SelStart = lngCursorPos
    End If
    
    Me.MousePointer = vbNormal
    
    If blnSetFocusOnTextbox Then
        rtfFormula.SetFocus
    ElseIf blnSetFocusOnChargeTextbox Then
        txtChargeState.SetFocus
    End If
    
    Exit Sub
    
StartIsotopicDistributionCalcsErrorHandler:
    GeneralErrorHandler "frmIsotopicDistribution|StartIsotopicDistributionCalcs", Err.Number, Err.Description
    
End Sub

Private Sub UpdateComparisonListStatus()
    lblComparisonListDataPoints = Trim(ComparisonListDataCount)
End Sub

Private Sub cboComparisonListPlotType_Click()
    PlotIsotopicDistribution True
End Sub

Private Sub cboPlotType_Click()
    PlotIsotopicDistribution True
End Sub

Private Sub chkAutoLabelPeaks_Click()
    PlotIsotopicDistribution True
End Sub

Private Sub chkPlotResults_Click()
    If cChkBox(chkPlotResults) Then PlotIsotopicDistribution True
End Sub

Private Sub cmdClose_Click()
    HideFormShowMain Me
End Sub

Private Sub cmdComparisonListClear_Click()
    ComparisonIonListClear
End Sub

Private Sub cmdComparisonListPaste_Click()
    ComparisonIonListPaste
End Sub

Private Sub cmdCompute_Click()
    StartIsotopicDistributionCalcs True
End Sub

Private Sub cmdCopy_Click()
    CopyResults
End Sub

Private Sub cmdShowPlot_Click()
    objSpectrum.ShowSpectrum
End Sub

Private Sub Form_Activate()
    PossiblyHideMainWindow
    
    mFormActivatedByUser = True
    
End Sub

Private Sub Form_Load()
    
    SizeAndCenterWindow Me, cWindowTopCenter, 11000, 5650
    
    If objSpectrum Is Nothing Then
        Set objSpectrum = New CWSpectrumDLL.Spectrum
    End If
    
    With objSpectrum
        .SetSeriesCount 2
        .SetSeriesPlotMode 1, pmLines, True
    End With
    mNeedToZoomOutFull = True
    
    mDelayUpdatingPlot = False
    
    ConvolutedMSDataCount = 0
    ReDim ConvolutedMSData(0)
    
    ComparisonListDataCount = 0
    ReDim ComparisonListXVals(0)
    ReDim ComparisonListYVals(0)
    
    PopulateComboBox cboPlotType, True, "Sticks to Zero|Gaussian Peaks", ipmGaussian    ' 15140
    PopulateComboBox cboComparisonListPlotType, True, "Sticks to Zero|Gaussian Peaks|Lines Between Points", ipmSticksToZero    ' 15145
    
    txtChargeState = "1"
    
    txtResults.Text = ""
    txtResults.Font = "Courier"
    
    PositionFormControls
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    QueryUnloadFormHandler Me, Cancel, UnloadMode
End Sub

Private Sub Form_Resize()
    PositionFormControls
End Sub

Private Sub lblComparisonListPlotColor_Click()
    SelectCustomColor Me, lblComparisonListPlotColor
    PlotIsotopicDistribution True
End Sub

Private Sub lblPlotColor_Click()
    SelectCustomColor Me, lblPlotColor
    PlotIsotopicDistribution True
End Sub

Private Sub objSpectrum_SpectrumFormRequestClose()
    ' The SpectrumForm was closed (actually, most likely just hidden)
    ' If we wanted to do anything special, we could do it here
End Sub

Private Sub rtfFormula_Change()
    Dim saveloc As Integer
        
    saveloc = rtfFormula.SelStart
    
    rtfFormula.TextRTF = objMwtWin.TextToRTF(rtfFormula.Text)
    rtfFormula.SelStart = saveloc
    
End Sub

Private Sub rtfFormula_GotFocus()
    SetMostRecentTextBoxValue rtfFormula.Text
    mLastControlWithFocus = lcwfFormula
End Sub

Private Sub rtfFormula_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        StartIsotopicDistributionCalcs True
        KeyAscii = 0
    End If
    
    ' Check the validity of the key using RTFBoxKeyPressHandler
    If KeyAscii <> 0 Then RTFBoxKeyPressHandler Me, rtfFormula, KeyAscii, False

End Sub

Private Sub txtChargeState_GotFocus()
    HighlightOnFocus txtChargeState
    mLastControlWithFocus = lcwfChargeState
End Sub

Private Sub txtChargeState_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtChargeState, KeyAscii, True, False
End Sub

Private Sub txtChargeState_LostFocus()
    ValidateTextboxValueLng txtChargeState, 0, 10000, 1
    If txtChargeState <> GetMostRecentTextBoxValue() Then
        StartIsotopicDistributionCalcs False, True
    End If
End Sub

Private Sub txtEffectiveResolution_GotFocus()
    HighlightOnFocus txtEffectiveResolution
    mLastControlWithFocus = lcwfEffectiveResolution
End Sub

Private Sub txtEffectiveResolution_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtEffectiveResolution, KeyAscii, True, False
End Sub

Private Sub txtEffectiveResolution_LostFocus()
    ValidateTextboxValueLng txtEffectiveResolution, 1, 1000000000#, 5000
    If txtEffectiveResolution <> GetMostRecentTextBoxValue() Then PlotIsotopicDistribution True
End Sub

Private Sub txtEffectiveResolutionMass_GotFocus()
    HighlightOnFocus txtEffectiveResolutionMass
    mLastControlWithFocus = lcwfEffectiveResolutionMass
End Sub

Private Sub txtEffectiveResolutionMass_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtEffectiveResolutionMass, KeyAscii, True, False
End Sub

Private Sub txtEffectiveResolutionMass_LostFocus()
    ValidateTextboxValueLng txtEffectiveResolutionMass, 1, 1000000000#, 1000
    If txtEffectiveResolutionMass <> GetMostRecentTextBoxValue() Then PlotIsotopicDistribution True
End Sub

Private Sub txtGaussianQualityFactor_GotFocus()
    HighlightOnFocus txtGaussianQualityFactor
    mLastControlWithFocus = lcwfGaussianQualityFactor
End Sub

Private Sub txtGaussianQualityFactor_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtGaussianQualityFactor, KeyAscii, True, False
End Sub

Private Sub txtGaussianQualityFactor_LostFocus()
    ValidateTextboxValueLng txtGaussianQualityFactor, 1, 75, 50
    If txtGaussianQualityFactor <> GetMostRecentTextBoxValue() Then PlotIsotopicDistribution True
End Sub

Private Sub txtResults_KeyPress(KeyAscii As Integer)
    
    Select Case KeyAscii
    Case 1          ' Ctrl+A -- select entire textbox
        txtResults.SelStart = 0
        txtResults.SelLength = Len(txtResults.Text)
        KeyAscii = 0
    Case 3          ' Copy is allowed
    Case 24, 3, 22  ' Cut and Paste are not allowed
        KeyAscii = 0
    Case 26         ' Ctrl+Z = Undo is not allowed
        KeyAscii = 0
    Case Else
        KeyAscii = 0
    End Select

End Sub

