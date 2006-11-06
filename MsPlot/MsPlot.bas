Attribute VB_Name = "MsPlotRoutines"
Option Explicit

Private Const NUM_TWIPS_PER_DIGIT = 90
Public Const MAX_DATA_SETS = 2              ' Up to 2 data sets may be graphed simultaneously (uses indices 0 and 1)

Public Const LowestValueForDoubleDataType = -1.79E+308      ' -3.4E+38
Public Const HighestValueForDoubleDataType = 1.79E+308      '  3.4E+38

Public Type usrAxisOptions
    Show As Boolean         ' Whether or not to show the axis
    ShowLabels As Boolean   ' whether or not to label axes
    MajorTicksToShow As Integer ' The number of major tick values to show
    ShowMinorTicks As Boolean   ' Whether or not to show minor ticks
    MinorTickMinimumPixelSep As Integer   ' Minimum spacing in pixels between minor ticks before showing them
    
    ShowGridLinesMajor As Boolean
    ShowTickLinesMinor As Boolean
    
End Type

Public Type usrPlotRangeDetails
    Val As Double
    Pos As Long
End Type

Public Type usrPlotRangeAxis
    ValStart As usrPlotRangeDetails
    ValEnd As usrPlotRangeDetails
    ValNegativeValueCorrectionOffset As Double
End Type

Public Type usrPlotRange
    x As usrPlotRangeAxis
    y As usrPlotRangeAxis
End Type

Public Type usrGaussianOptions
    ResolvingPower As Long              ' Effective resolution  (M / delta M)
    XValueOfSpecification As Single     ' X Value where effective resolution applies
End Type

Public Type usrPlotDataOptions
    
    PlotTypeCode As Integer     ' Whether the plot is a stick plot (0) or line-between-points plot (1)
    GaussianConversion As usrGaussianOptions
    
    '  Zoom options
    ZoomOutFull As Boolean      ' When true, the data is shown fully zoomed out, ignoring the XVal and YVal Start and End values
    AutoScaleY As Boolean       ' When true, will zoom the x data as requested but always keep the y data at full scale
    FixYAxisMinimumAtZero As Boolean ' Only used if AutoScaleY = True; keeps the minimum y scaling at 0 at all times (i.e. makes YValStart = 0)
            
    ' Scaling options
    Scaling As usrPlotRange     ' Note that I do not use ValNegativeValueCorrectionOffset in Scaling, only in DataLimits
    DataLimits(MAX_DATA_SETS) As usrPlotRange      ' Records the largest and smallest values of the given data set
    SortData As Boolean         ' Whether or not to sort the data by x value before plotting
    
    ' The following define the size of the spectrum in the parent frame
    PlotLeft As Integer         ' In Twips (VB units)
    PlotTop As Integer          ' In Twips (VB units)
    PlotWidth As Integer        ' In Twips (VB units)
    PlotHeight As Integer       ' In Twips (VB units)
    PlotLeftLargeNumberOffset As Integer    ' In Twips - used to offset the left of the graph to the right a little for very large or very small numbers

    XAxis As usrAxisOptions
    YAxis As usrAxisOptions
    
    MajorTicksLoaded As Long    ' There is a label and gridline loaded for each major tick loaded
    MinorTicksLoaded As Long

    boolLongOperationsRequired As Boolean       ' Set to true when long operations are encountered, thus requiring an hourglass cursor
    
    ' Labeling options
    ShowDataPointLabels As Boolean       ' Whether to label or not
    LabelsToShow As Integer     ' Number of labels to show
    IndexToHighlight As Long    ' Index of the x,y data pair to highlight (if any)
    HighlightColor As Long      ' default vbRed

    LinesLoadedCount As Long    ' Records the number of lines and labels that have been dynamically loaded
End Type

Public Type usrXYData
    XVal As Double          ' Note: Single data type is -3.402823E38 to 3.402823E38
    YVal As Double
End Type

Public Type usrXYDataSet
    XYDataList() As usrXYData                 ' Index 1 to XYDataListCount
    XYDataListCount As Long
    XYDataListCountDimmed As Long
End Type

Private Sub CheckForLongOperation(ThisForm As Form, lngSecondsElapsedAtStart As Long, PlotOptions As usrPlotDataOptions)
    ' Checks to see if the current Timer() value (sec elapsed since midnight) is
    '  greater than lngSecondsElapsedAtStart.  If it is, it turns boolSavedCheckBit on
    '  and sets the pointer to an hourglass
    ' The boolSavedCheckBit value is saved for future calls to the sub so that the hourglass
    '  will be activated immediately on future calls
    
    If Timer() - lngSecondsElapsedAtStart > 1 Or PlotOptions.boolLongOperationsRequired Then
        PlotOptions.boolLongOperationsRequired = True
        ThisForm.MousePointer = vbHourglass
    End If
    
End Sub

Public Function ConstructFormatString(dblThisValue As Double, Optional ByRef intDigitsInFormattedValue As Integer) As String
    ' Examines dblThisValue and constructs a format string based on its magnitude
    ' For example, dblThisValue = 1234 will return "0"
    '              dblThisValue = 2.4323 will return "0.0000"
    '
    ' In addition, returns the length of the string representation of dblThisValue using the determined format string
    Dim lngExponentValue As Long, intDigitsInLabel As Integer
    Dim strWork As String, strFormatString As String
    
    ' Determine width of label to use and construct formatting string for labels
    ' First, find the exponent of dblThisValue
    strWork = Format(dblThisValue, "0E+000")
    lngExponentValue = CIntSafe(Right(strWork, 4))
    
    ' Determine number of digits in dblThisValue, rounded according to lngExponentVal
    If lngExponentValue >= 0 Then
        intDigitsInLabel = 0
        strFormatString = "0"
    Else
        ' Add 1 for the decimal point
        intDigitsInLabel = -lngExponentValue + 1
        strFormatString = "0." & String(-lngExponentValue, "0")
    End If

    intDigitsInFormattedValue = Len(Format(dblThisValue, strFormatString))
    
    ConstructFormatString = strFormatString
End Function

Public Sub CheckDynamicXYData(ByRef ThisXYDataList() As usrXYData, ThisXYDataListCount As Long, ByRef ThisXYDataListCountDimmed As Long, Optional lngIncrement As Long = 100)
    If ThisXYDataListCount > ThisXYDataListCountDimmed Then
        ThisXYDataListCountDimmed = ThisXYDataListCountDimmed + lngIncrement
        If ThisXYDataListCountDimmed < ThisXYDataListCount Then
            ThisXYDataListCountDimmed = ThisXYDataListCount
        End If
        ReDim Preserve ThisXYDataList(ThisXYDataListCountDimmed)
    End If
End Sub

Public Function ConvertStickDataToGaussian(ThisForm As Form, ThisXYDataSet As usrXYDataSet, PlotOptions As usrPlotDataOptions, intDataSetIndex As Integer) As usrXYDataSet
    Dim XYDataPointerArray() As Long, XYDataPointerArrayCount As Long
    Dim lngDataIndex As Long, lngMidPointIndex As Long
    Dim lngStickIndex As Long, DeltaX As Double
    
    Dim dblXValRange As Double, dblXValWindowRange As Double, dblRangeWork As Double
    Dim dblMinimalXValOfWindow As Double, dblMinimalXValSpacing As Double
    
    Dim dblXOffSet As Double, sigma As Double
    Dim intDigitsToRoundTo As Integer, lngExponentValue As Long
    Dim lngSecElapsedAtStart As Long
    
    Dim XYSummation() As usrXYData, XYSummationCount As Long, XYSummationCountDimmed As Long
    Dim lngSummationIndex As Long, lngMinimalSummationIndex As Long, lngOffsetIndex As Long
    Dim DataToAdd() As usrXYData, lngDataToAddCount As Long, boolAppendNewData As Boolean
    Dim boolSequentialValueFound As Boolean
    Dim ThisDataPoint As usrXYData
    
    Const NUM_CHUNKS = 100
    
    If ThisXYDataSet.XYDataListCount = 0 Then Exit Function
    
    lngSecElapsedAtStart = Timer()
    
    If PlotOptions.GaussianConversion.ResolvingPower < 1 Then
        PlotOptions.GaussianConversion.ResolvingPower = 1
    End If
    
    XYSummationCount = 0
    XYSummationCountDimmed = 100
    ReDim XYSummation(XYSummationCountDimmed)
    
    With ThisXYDataSet
        ' Initialize the Pointer Array
        XYDataPointerArrayCount = .XYDataListCount
        ReDim XYDataPointerArray(XYDataPointerArrayCount)
        For lngDataIndex = 1 To XYDataPointerArrayCount
            XYDataPointerArray(lngDataIndex) = lngDataIndex
        Next lngDataIndex
        
        ' See if data is sorted; if not, sort it
        ' Force .SortData to be True
        PlotOptions.SortData = True
        CheckDataLimitsAndSortData ThisForm, .XYDataList(), XYDataPointerArray(), XYDataPointerArrayCount, PlotOptions, intDataSetIndex
    
        PlotOptions.DataLimits(intDataSetIndex).x.ValStart.Val = .XYDataList(XYDataPointerArray(1)).XVal
        PlotOptions.DataLimits(intDataSetIndex).x.ValEnd.Val = .XYDataList(XYDataPointerArray(XYDataPointerArrayCount)).XVal
    End With
    
    With PlotOptions
        dblXValRange = .DataLimits(intDataSetIndex).x.ValEnd.Val - .DataLimits(intDataSetIndex).x.ValStart.Val
        
        ' Set DeltaX using .ResolvingPower and .XValueOfSpecification
        ' Do not allow the DeltaX to be so small that the total points required > 100,000
        DeltaX = .GaussianConversion.XValueOfSpecification / .GaussianConversion.ResolvingPower / 20
        ' Make sure DeltaX is a reasonable number
        DeltaX = RoundToMultipleOf10(DeltaX)
        If DeltaX = 0 Then DeltaX = 1
        
        ' Set the Window Range to 1/10 the magnitude of the midpoint x value
        dblRangeWork = .DataLimits(intDataSetIndex).x.ValStart.Val + dblXValRange / 2
        dblRangeWork = RoundToMultipleOf10(dblRangeWork, lngExponentValue)
        
        sigma = (.GaussianConversion.XValueOfSpecification / .GaussianConversion.ResolvingPower) / Sqr(5.54)
        
        ' Set the window range (the xvalue window width range) to calculate the Gaussian representation for each data point
        ' The width at the base of a peak is 4 sigma
        ' Use a width of 2 * 6 sigma
        dblXValWindowRange = 2 * 6 * sigma
        
        If dblXValWindowRange / DeltaX > 50000 Then
            DeltaX = dblXValWindowRange / 50000
        End If
        
        lngDataToAddCount = CLng(dblXValWindowRange / DeltaX)
        ' Make sure lngDataToAddCount is odd
        If CSng(lngDataToAddCount) / 2! = Round(CSng(lngDataToAddCount) / 2!, 0) Then
            lngDataToAddCount = lngDataToAddCount + 1
        End If
        ReDim DataToAdd(lngDataToAddCount)
        lngMidPointIndex = (lngDataToAddCount + 1) / 2
    End With
    
    With ThisXYDataSet
'''        frmProgress.InitializeForm "Creating Gaussian Representation", 0, XYDataPointerArrayCount
        For lngStickIndex = 1 To XYDataPointerArrayCount
'''            frmProgress.UpdateProgressBar lngStickIndex
            If lngStickIndex Mod 50 = 0 Then
                CheckForLongOperation ThisForm, lngSecElapsedAtStart, PlotOptions
            End If
            
            ' Search through XYSummation to determine the index of the smallest XValue with which
            '   data in DataToAdd could be combined
            lngMinimalSummationIndex = 1
            dblMinimalXValOfWindow = .XYDataList(XYDataPointerArray(lngStickIndex)).XVal - (lngMidPointIndex - 1) * DeltaX
            If dblMinimalXValOfWindow > XYSummation(XYSummationCount).XVal Then
                lngMinimalSummationIndex = XYSummationCount
            Else
                For lngSummationIndex = 1 To XYSummationCount
                    If XYSummation(lngSummationIndex).XVal >= dblMinimalXValOfWindow Then
                        lngMinimalSummationIndex = lngSummationIndex - 1
                        If lngMinimalSummationIndex < 1 Then lngMinimalSummationIndex = 1
                        Exit For
                    End If
                Next lngSummationIndex
                If lngSummationIndex > XYSummationCount Then
                    lngMinimalSummationIndex = XYSummationCount
                End If
            End If
            
            ' Construct the Gaussian representation for this Data Point
            ThisDataPoint = .XYDataList(XYDataPointerArray(lngStickIndex))
            
            ' Round ThisDataPoint.XVal to the nearest DeltaX
            ' If .XVal is not an even multiple of DeltaX then bump up .XVal until it is
            ThisDataPoint.XVal = RoundToEvenMultiple(ThisDataPoint.XVal, DeltaX, True)
            
            For lngDataIndex = 1 To lngDataToAddCount
                ' Equation for Gaussian is: Amplitude * Exp[ -(x - mu)^2 / (2*sigma^2) ]
                '        Use lngDataIndex, .YVal, and DeltaX
                dblXOffSet = (lngMidPointIndex - lngDataIndex) * DeltaX
                DataToAdd(lngDataIndex).XVal = ThisDataPoint.XVal - dblXOffSet
                DataToAdd(lngDataIndex).YVal = ThisDataPoint.YVal * Exp(-(dblXOffSet) ^ 2 / (2 * sigma ^ 2))
            Next lngDataIndex
            
            ' Now merge DataToAdd into XYSummation
            ' XValues in DataToAdd and those in XYSummation have the same DeltaX value
            ' The XValues in DataToAdd might overlap partially with those in XYSummation
            
            lngDataIndex = 1
            ' First, see if the first XValue in DataToAdd is larger than the last XValue in XYSummation
            If DataToAdd(lngDataIndex).XVal > XYSummation(XYSummationCount).XVal Then
                boolAppendNewData = True
            Else
                boolAppendNewData = False
                ' Step through XYSummation() starting at lngMinimalSummationIndex, looking for
                '   the index to start combining data at
                For lngSummationIndex = lngMinimalSummationIndex To XYSummationCount
                    If DataToAdd(lngDataIndex).XVal = XYSummation(lngSummationIndex).XVal Or DataToAdd(lngDataIndex).XVal < XYSummation(lngSummationIndex).XVal Then
                        '''Debug.Assert XYSummation(lngSummationIndex).XVal = DataToAdd(lngDataIndex).XVal
                        ' Within Tolerance; start combining the values here
                        Do While lngSummationIndex <= XYSummationCount
                            XYSummation(lngSummationIndex).YVal = XYSummation(lngSummationIndex).YVal + DataToAdd(lngDataIndex).YVal
                            lngSummationIndex = lngSummationIndex + 1
                            lngDataIndex = lngDataIndex + 1
                            If lngDataIndex > lngDataToAddCount Then
                                ' Successfully combined all of the data
                                Exit Do
                            End If
                        Loop
                        If lngDataIndex <= lngDataToAddCount Then
                            ' Data still remains to be added
                            boolAppendNewData = True
                        End If
                        Exit For
                    End If
                Next lngSummationIndex
            End If
            
            If boolAppendNewData = True Then
                CheckDynamicXYData XYSummation(), XYSummationCount + lngDataToAddCount - lngDataIndex + 1, XYSummationCountDimmed
                Do While lngDataIndex <= lngDataToAddCount
                    ThisDataPoint = DataToAdd(lngDataIndex)
                    XYSummationCount = XYSummationCount + 1
                    XYSummation(XYSummationCount) = ThisDataPoint
                    lngDataIndex = lngDataIndex + 1
                Loop
            End If
            
        Next lngStickIndex
        
    End With
    
    ' Step through XYSummation and remove areas of sequential equivalent values
    lngSummationIndex = 1
    RoundToMultipleOf10 DeltaX, lngExponentValue
    If lngExponentValue < 0 Then
        intDigitsToRoundTo = Abs(lngExponentValue)
    Else
        intDigitsToRoundTo = 0
    End If
    
    ' Assure there is a data point at each 1% point along x range
    ' Probably need to add data, but may need to remove some
    dblMinimalXValSpacing = dblXValRange / 100
    
    Do While lngSummationIndex <= XYSummationCount - 1
        If XYSummation(lngSummationIndex).XVal + dblMinimalXValSpacing < XYSummation(lngSummationIndex + 1).XVal Then
            ' Need to insert a data point
            XYSummationCount = XYSummationCount + 1
            CheckDynamicXYData XYSummation, XYSummationCount, XYSummationCountDimmed
            For lngDataIndex = XYSummationCount To lngSummationIndex + 2 Step -1
                XYSummation(lngDataIndex) = XYSummation(lngDataIndex - 1)
            Next lngDataIndex
            XYSummation(lngSummationIndex + 1).XVal = XYSummation(lngSummationIndex + 1).XVal + dblMinimalXValSpacing
            XYSummation(lngSummationIndex + 1).YVal = (XYSummation(lngSummationIndex + 1).YVal + XYSummation(lngSummationIndex + 2).YVal) / 2
        End If
        lngSummationIndex = lngSummationIndex + 1
    Loop
    
    HideProgressForm

    ' Reset the boolLongOperationsRequired bit
    PlotOptions.boolLongOperationsRequired = False

    ' ReDim XYSummation to XYSummationCount since DrawPlot assumes this is the case
    XYSummationCountDimmed = XYSummationCount
    ReDim Preserve XYSummation(XYSummationCountDimmed)
    
    ' Assign data in XYSummation to the function so that it gets returned
    ConvertStickDataToGaussian.XYDataList = XYSummation
    ConvertStickDataToGaussian.XYDataListCount = XYSummationCount
    ConvertStickDataToGaussian.XYDataListCountDimmed = XYSummationCountDimmed

End Function

Private Sub CheckDataLimitsAndSortData(ThisForm As Form, ThisXYData() As usrXYData, XYDataPointerArray() As Long, XYDataPointerArrayCount As Long, PlotOptions As usrPlotDataOptions, intDataSetIndex As Integer)
    Dim boolNeedsSorting As Boolean, lngIndex As Long
    Dim dblMaximumIntensity As Double, dblMinimumIntensity As Double
    Dim dblXYDataPoint As Double
    
    
    If PlotOptions.DataLimits(intDataSetIndex).y.ValStart.Val = 0 And PlotOptions.DataLimits(intDataSetIndex).y.ValEnd.Val = 0 Then
        ' Data Limits not defined
        ' Figure out what they are and sort data if necessary
        
        ' Find the Y scale data limits
        ' At the same time, see if the data is sorted
        dblMaximumIntensity = LowestValueForDoubleDataType
        dblMinimumIntensity = HighestValueForDoubleDataType
        boolNeedsSorting = False
        For lngIndex = 1 To XYDataPointerArrayCount
            dblXYDataPoint = ThisXYData(XYDataPointerArray(lngIndex)).YVal
            If dblXYDataPoint > dblMaximumIntensity Then dblMaximumIntensity = dblXYDataPoint
            If dblXYDataPoint < dblMinimumIntensity Then dblMinimumIntensity = dblXYDataPoint
            If lngIndex < XYDataPointerArrayCount Then
                If ThisXYData(XYDataPointerArray(lngIndex)).XVal > ThisXYData(XYDataPointerArray(lngIndex + 1)).XVal Then
                    boolNeedsSorting = True
                End If
            End If
        Next lngIndex
        PlotOptions.DataLimits(intDataSetIndex).y.ValStart.Val = dblMinimumIntensity
        PlotOptions.DataLimits(intDataSetIndex).y.ValEnd.Val = dblMaximumIntensity
        
        If boolNeedsSorting And PlotOptions.SortData Then
            ' Need to sort data
            ThisForm.MousePointer = vbHourglass
            SortXYData ThisXYData(), XYDataPointerArray(), XYDataPointerArrayCount, False
        End If
    End If

End Sub

Public Sub DrawPlot(ThisForm As Form, PlotOptions As usrPlotDataOptions, ThisXYDataArray() As usrXYDataSet, ByRef PlotRange() As usrPlotRange, intDataSetsLoaded As Integer)
    ' Draw a graphical representation of a list of x,y data pairs in one or more data sets (stored in ThisXYDataArray)
    ' Assumes the x,y data point array is 1-based (i.e. the first data point is in index 1
    
    Dim intDataSetIndex As Integer
    Dim lngIndex As Long, lngLineIndex As Long
    Dim PlotBottom As Long, lngLeftOffset As Long
    Dim lngSecElapsedAtStart As Long
    Dim intKeepEveryXthPoint As Integer, lngXYDataToCountTrack As Long
    
    Const MaxLinesCount = 32000
    
    Dim XYDataToPlot() As usrXYData, XYDataToPlotCount As Long
    
    Dim XYDataPointerArray() As Long
    Dim XYDataPointerArrayCount As Long, XYDataIndex As Long
    
    Dim HighlightIndex As Long
    Dim dblPreviousMinimum As Double
    Dim lngChunkSize As Long, lngMinimumValIndex As Long, lngMaximumValIndex As Long
    Dim boolNeedsSorting As Boolean
    
    Dim StartXVal As Double, XValRange As Double, DeltaXScaler As Double
    Dim StartXValIndex As Long, EndXValIndex As Long
    Dim strWork As String, intExponentValue As Integer, dblWork As Double

    Dim LastLabelTwip As Long, TwipsBetweenLabels As Long
    Dim intDynamicObjectOffset As Integer, lngDataSetLineColor As Long
    Dim dblMinimumIntensity As Double, dblMaximumIntensity As Double
    
    ' Need to determine the correct scaling values if autoscaling the y-axis
    ' Only works if the data is sorted in the x direction
    If PlotOptions.AutoScaleY Or PlotOptions.FixYAxisMinimumAtZero Then
        ' Find the minimum and maximum y intensities for all data sets within the range of x data being shown
        dblMaximumIntensity = LowestValueForDoubleDataType
        dblMinimumIntensity = HighestValueForDoubleDataType
        For intDataSetIndex = 0 To intDataSetsLoaded - 1
            With ThisXYDataArray(intDataSetIndex)
                ' Step through .XYDataList and find index of the start x value
                For lngIndex = 1 To .XYDataListCount
                    If .XYDataList(lngIndex).XVal >= PlotOptions.Scaling.x.ValStart.Val Then
                        StartXValIndex = lngIndex
                        Exit For
                    End If
                Next lngIndex
            
                ' Step through .XYDataList and find index of the end x value
                For lngIndex = .XYDataListCount To 1 Step -1
                    If .XYDataList(lngIndex).XVal <= PlotOptions.Scaling.x.ValEnd.Val Then
                        EndXValIndex = lngIndex
                        Exit For
                    End If
                Next lngIndex
                
                For lngIndex = StartXValIndex To EndXValIndex
                    dblWork = .XYDataList(lngIndex).YVal
                    If dblWork > dblMaximumIntensity Then
                        dblMaximumIntensity = dblWork
                    End If
                    If dblWork < dblMinimumIntensity Then
                        dblMinimumIntensity = dblWork
                    End If
                Next lngIndex
            End With
        Next intDataSetIndex
        
        If PlotOptions.FixYAxisMinimumAtZero Then
            ' Only fix Y axis range at zero, and only fix if .FixYAxisMinimumAtZero is true
            PlotOptions.Scaling.y.ValStart.Val = 0
        Else
            PlotOptions.Scaling.y.ValStart.Val = dblMinimumIntensity    ' + ThisAxisRange.ValNegativeValueCorrectionOffset
        End If
        
        PlotOptions.Scaling.y.ValEnd.Val = dblMaximumIntensity                ' + ThisAxisRange.ValNegativeValueCorrectionOffset
    End If

    intDynamicObjectOffset = 0
    For intDataSetIndex = 0 To intDataSetsLoaded - 1
        Select Case intDataSetIndex
        Case 0: lngDataSetLineColor = vbBlack
        Case 1: lngDataSetLineColor = vbRed
        Case 2: lngDataSetLineColor = vbGreen
        Case Else: lngDataSetLineColor = vbMagenta
        End Select
        
        If ThisXYDataArray(intDataSetIndex).XYDataListCount > 0 Then
            
            lngSecElapsedAtStart = Timer()
            
            ' Initialize the pointer array
            XYDataPointerArrayCount = ThisXYDataArray(intDataSetIndex).XYDataListCount
            ReDim XYDataPointerArray(XYDataPointerArrayCount)
            For lngIndex = 1 To XYDataPointerArrayCount
                XYDataPointerArray(lngIndex) = lngIndex
            Next lngIndex
            
            ' See if data is sorted; if not, sort it
            CheckDataLimitsAndSortData ThisForm, ThisXYDataArray(intDataSetIndex).XYDataList, XYDataPointerArray(), XYDataPointerArrayCount, PlotOptions, intDataSetIndex
                
            With ThisXYDataArray(intDataSetIndex)
                
                ' Determine the location in the parent frame of the bottom of the Plot
                PlotBottom = PlotOptions.PlotTop + PlotOptions.PlotHeight
                
                StartXValIndex = 1
                EndXValIndex = .XYDataListCount
                
                ' Record the X scale data limits
                PlotOptions.DataLimits(intDataSetIndex).x.ValStart.Val = .XYDataList(XYDataPointerArray(StartXValIndex)).XVal
                PlotOptions.DataLimits(intDataSetIndex).x.ValEnd.Val = .XYDataList(XYDataPointerArray(EndXValIndex)).XVal
                
                ' Initialize .Scaling.x.ValStart.Val and .Scaling.y.ValStart.Val if necessary
                If PlotOptions.Scaling.x.ValStart.Val = 0 And PlotOptions.Scaling.x.ValEnd.Val = 0 Then
                    PlotOptions.Scaling.x.ValStart.Val = .XYDataList(XYDataPointerArray(StartXValIndex)).XVal
                    PlotOptions.Scaling.x.ValEnd.Val = .XYDataList(XYDataPointerArray(EndXValIndex)).XVal
                End If
                
                ' Make sure .Scaling.X.ValStart.Val < .Scaling.X.ValEnd
                If PlotOptions.Scaling.x.ValStart.Val > PlotOptions.Scaling.x.ValEnd.Val Then
                    SwapValues PlotOptions.Scaling.x.ValStart.Val, PlotOptions.Scaling.x.ValEnd.Val
                End If
                
                If Not PlotOptions.ZoomOutFull Then
                    ' Step through .XYDataList and find index of the start x value
                    For lngIndex = 1 To .XYDataListCount
                        If .XYDataList(XYDataPointerArray(lngIndex)).XVal >= PlotOptions.Scaling.x.ValStart.Val Then
                            StartXValIndex = lngIndex
                            Exit For
                        End If
                    Next lngIndex
                
                    ' Step through .XYDataList and find index of the end x value
                    For lngIndex = .XYDataListCount To 1 Step -1
                        If .XYDataList(XYDataPointerArray(lngIndex)).XVal <= PlotOptions.Scaling.x.ValEnd.Val Then
                            EndXValIndex = lngIndex
                            Exit For
                        End If
                    Next lngIndex
                End If
                
                ' Make sure StartValIndex <= EndValIndex
                If StartXValIndex > EndXValIndex Then
                    SwapValues StartXValIndex, EndXValIndex
                End If
                
                ' Check to see if Mouse Pointer should be changed to hourglass
                CheckForLongOperation ThisForm, lngSecElapsedAtStart, PlotOptions
                
                ' Copy the data into XYDataToPlot, reindexing to start at 1 and going to XYDataToPlotCount
                ' Although this uses more memory because it duplicates the data, I soon replace the
                ' raw data values with location positions in twips
                '
                ' In addition, if there is far more data than could be possibly plotted,
                ' I throw away every xth data point, though with a twist
                
                XYDataToPlotCount = EndXValIndex - StartXValIndex + 1
                
                If XYDataToPlotCount > PlotOptions.PlotWidth / 10 Then
                    ' Throw away some of the data:  Note that CIntSafe will round 2.5 up to 3
                    intKeepEveryXthPoint = CIntSafeDbl(XYDataToPlotCount / (PlotOptions.PlotWidth / 10))
                    lngChunkSize = intKeepEveryXthPoint * 2
                Else
                    intKeepEveryXthPoint = 1
                End If
                
                ReDim XYDataToPlot(CLngRoundUp(XYDataToPlotCount / intKeepEveryXthPoint) + 10)
                
                If intKeepEveryXthPoint = 1 Then
                    lngXYDataToCountTrack = 0
                    For lngIndex = StartXValIndex To EndXValIndex
                        lngXYDataToCountTrack = lngXYDataToCountTrack + 1
                        XYDataToPlot(lngXYDataToCountTrack) = .XYDataList(XYDataPointerArray(lngIndex))
                        If XYDataPointerArray(lngIndex) = PlotOptions.IndexToHighlight Then
                            HighlightIndex = lngXYDataToCountTrack
                        End If
                    Next lngIndex
                Else
                    ' Step through the data examining chunks of numbers twice the length of intKeepEveryXthPoint
                    ' Find the minimum and maximum value in each chunk of numbers
                    ' Store these values in the array to keep (minimum followed by maximum)
                    ' Swap the stored order if both numbers are less than the previous two numbers
                    
                    ' Store the first value of .xydatalist() in the output array
                    lngXYDataToCountTrack = 1
                    XYDataToPlot(lngXYDataToCountTrack) = .XYDataList(XYDataPointerArray(StartXValIndex))
                    If XYDataPointerArray(StartXValIndex) = PlotOptions.IndexToHighlight Then
                        HighlightIndex = lngXYDataToCountTrack
                    End If
                    
                    dblPreviousMinimum = LowestValueForDoubleDataType
                    For lngIndex = StartXValIndex + 1 To EndXValIndex - 1 Step lngChunkSize
                        
                        FindMinimumAndMaximum lngMinimumValIndex, lngMaximumValIndex, .XYDataList(), XYDataPointerArray(), lngIndex, lngIndex + lngChunkSize
                        
                        ' Check if the maximum value of this pair of points is less than the minimum value of the previous pair
                        ' If it is, the y values of the two points should be exchanged
                        If .XYDataList(XYDataPointerArray(lngMaximumValIndex)).YVal < dblPreviousMinimum Then
                            ' Update dblPreviousMinimum
                            dblPreviousMinimum = .XYDataList(XYDataPointerArray(lngMinimumValIndex)).YVal
                            ' Swap minimum and maximum so that maximum gets saved to array first
                            SwapValues lngMinimumValIndex, lngMaximumValIndex
                        Else
                            ' Update dblPreviousMinimum
                            dblPreviousMinimum = .XYDataList(XYDataPointerArray(lngMinimumValIndex)).YVal
                        End If
                            
                        lngXYDataToCountTrack = lngXYDataToCountTrack + 1
                          XYDataToPlot(lngXYDataToCountTrack) = .XYDataList(XYDataPointerArray(lngMinimumValIndex))
                          If XYDataPointerArray(lngMinimumValIndex) = PlotOptions.IndexToHighlight Then
                            HighlightIndex = lngXYDataToCountTrack
                          End If
                        
                        lngXYDataToCountTrack = lngXYDataToCountTrack + 1
                          XYDataToPlot(lngXYDataToCountTrack) = .XYDataList(XYDataPointerArray(lngMaximumValIndex))
                          If XYDataPointerArray(lngMaximumValIndex) = PlotOptions.IndexToHighlight Then
                            HighlightIndex = lngXYDataToCountTrack
                          End If
                    
                        ' Now check to see if the .XVal of the first value is greater than the first
                        ' If so, swap the .XVals
                        If XYDataToPlot(lngXYDataToCountTrack - 1).XVal > XYDataToPlot(lngXYDataToCountTrack).XVal Then
                            SwapValues XYDataToPlot(lngXYDataToCountTrack - 1).XVal, XYDataToPlot(lngXYDataToCountTrack).XVal
                        End If
                    Next lngIndex
                    
                    ' Store the last value of .xydatalist() in the output array
                    lngXYDataToCountTrack = lngXYDataToCountTrack + 1
                    XYDataToPlot(lngXYDataToCountTrack) = .XYDataList(XYDataPointerArray(EndXValIndex))
                    If XYDataPointerArray(EndXValIndex) = PlotOptions.IndexToHighlight Then
                        HighlightIndex = lngXYDataToCountTrack
                    End If
                
                End If
                
                ' Set XYDataToPlotCount to the lngXYDataToCountTrack value resulting from the copying
                XYDataToPlotCount = lngXYDataToCountTrack
            
            End With
        
            ' Scale the data vertically according to PlotOptions.height
            ' i.e., replace the actual y values with locations in twips for where the data point belongs on the graph
            ' The new value will range from 0 to PlotOptions.Height
            ' Note that PlotOptions.PlotLeftLargeNumberOffset is computed in ScaleData for the Y axis
            ScaleData PlotOptions, XYDataToPlot(), XYDataToPlotCount, PlotRange(intDataSetIndex).y, PlotOptions.Scaling.y, False, PlotOptions.AutoScaleY
            
            ' Now scale the data to twips, ranging 0 to .PlotWidth and 0 to .PlotHeight
            ' X axis
            ScaleData PlotOptions, XYDataToPlot(), XYDataToPlotCount, PlotRange(intDataSetIndex).x, PlotOptions.Scaling.x, True, False
            
            ' Load lines and labels for each XVal
            If PlotOptions.LinesLoadedCount = 0 Then
                ' Only initialize to 1 the first time this sub is called
                PlotOptions.LinesLoadedCount = 1
            End If
            
            ' Limit the total points shown if greater than MaxLinesCount
            If XYDataToPlotCount > MaxLinesCount Then
                ' This code should not be reached since extra data should have been thrown away above
                Debug.Assert False
                XYDataToPlotCount = MaxLinesCount
            End If
            
            ' Load dynamic plot objects as needed
            LoadDynamicPlotObjects ThisForm, PlotOptions, intDynamicObjectOffset + XYDataToPlotCount
            
            ' Must re-initialize the pointer array
            XYDataPointerArrayCount = XYDataToPlotCount
            ReDim XYDataPointerArray(XYDataToPlotCount)
            For lngIndex = 1 To XYDataToPlotCount
                XYDataPointerArray(lngIndex) = lngIndex
            Next lngIndex
            
            If PlotOptions.PlotTypeCode = 0 Then
                ' Plot the data as sticks to zero
                If PlotOptions.ShowDataPointLabels And PlotOptions.LabelsToShow > 0 Then
                    ' Sort the data by YVal, draw in order of YVal, and only label top specified peaks
                    SortXYData XYDataToPlot, XYDataPointerArray, XYDataPointerArrayCount, True
                End If
            Else
                ' Plot the data as lines between points
                ' Plot in order of XVal, thus, don't need to sort
                PlotOptions.ShowDataPointLabels = False
            End If
            
            ' Check to see if Mouse Pointer should be changed to hourglass
            CheckForLongOperation ThisForm, lngSecElapsedAtStart, PlotOptions
            
            ' Label the axes and add ticks and gridlines
            FormatAxes ThisForm, PlotOptions, PlotBottom, PlotRange(intDataSetIndex)
            
            ' Position the lines and labels
            LastLabelTwip = 0
                
            lngLeftOffset = PlotOptions.PlotLeft + PlotOptions.PlotLeftLargeNumberOffset
            
            For lngIndex = 1 To XYDataToPlotCount
                With ThisForm
                    If lngIndex Mod 50 = 0 Then
                        ' Check to see if Mouse Pointer should be changed to hourglass
                        CheckForLongOperation ThisForm, lngSecElapsedAtStart, PlotOptions
                    End If
                    
                    XYDataIndex = XYDataPointerArray(lngIndex)
                    lngLineIndex = lngIndex + intDynamicObjectOffset
                    
                    .linData(lngLineIndex).Visible = True
        
                    If PlotOptions.PlotTypeCode = 0 Then
                        ' Plot the data as sticks to zero
                        
                        .linData(lngLineIndex).x1 = lngLeftOffset + XYDataToPlot(XYDataIndex).XVal
                        .linData(lngLineIndex).x2 = .linData(lngLineIndex).x1
                        .linData(lngLineIndex).y1 = PlotBottom
                        .linData(lngLineIndex).y2 = PlotBottom - XYDataToPlot(XYDataIndex).YVal
                        If HighlightIndex = XYDataIndex Then
                            .linData(lngLineIndex).BorderColor = PlotOptions.HighlightColor
                        Else
                            .linData(lngLineIndex).BorderColor = lngDataSetLineColor
                        End If
                    Else
                        ' Plot the data as lines between points
                        If lngIndex < XYDataToPlotCount Then
                            .linData(lngLineIndex).x1 = lngLeftOffset + XYDataToPlot(XYDataIndex).XVal
                            .linData(lngLineIndex).x2 = lngLeftOffset + XYDataToPlot(XYDataIndex + 1).XVal
                            .linData(lngLineIndex).y1 = PlotBottom - XYDataToPlot(XYDataIndex).YVal
                            .linData(lngLineIndex).y2 = PlotBottom - XYDataToPlot(XYDataIndex + 1).YVal
                            .linData(lngLineIndex).BorderColor = lngDataSetLineColor
                        Else
                            .linData(lngLineIndex).Visible = False
                        End If
                    End If
                End With
                
                With ThisXYDataArray(intDataSetIndex)
                    If PlotOptions.ShowDataPointLabels Then
                        
                        If XYDataIndex = HighlightIndex Or lngIndex <= PlotOptions.LabelsToShow Then
                            LastLabelTwip = ThisForm.lblPlotIntensity(lngLineIndex).Left
                            ThisForm.lblPlotIntensity(lngLineIndex).Visible = True
                            ThisForm.lblPlotIntensity(lngLineIndex).Tag = "Visible"
                            If XYDataIndex = HighlightIndex Then
                                ' Include charge in label
                                ThisForm.lblPlotIntensity(lngLineIndex).Caption = Format(.XYDataList(StartXValIndex + XYDataIndex - 1).XVal, "0.0000")
                                ThisForm.lblPlotIntensity(lngLineIndex).Width = 1200
                            Else
                                ThisForm.lblPlotIntensity(lngLineIndex).Caption = Format(.XYDataList(StartXValIndex + XYDataIndex - 1).XVal, "0.0000")
                                ThisForm.lblPlotIntensity(lngLineIndex).Width = 800
                            End If
                            ThisForm.lblPlotIntensity(lngLineIndex).ToolTipText = Format(.XYDataList(StartXValIndex + XYDataIndex - 1).XVal, "0.0000")
                            ThisForm.lblPlotIntensity(lngLineIndex).Height = 200
                            ThisForm.lblPlotIntensity(lngLineIndex).Top = ThisForm.linData(lngLineIndex).y2 - ThisForm.lblPlotIntensity(lngLineIndex).Height
                            ThisForm.lblPlotIntensity(lngLineIndex).Left = ThisForm.linData(lngLineIndex).x1 - 200
                        Else
                            ThisForm.lblPlotIntensity(lngLineIndex).Visible = True
                            ThisForm.lblPlotIntensity(lngLineIndex).Tag = "Hidden"
                            ThisForm.lblPlotIntensity(lngLineIndex).Caption = ""
                            ThisForm.lblPlotIntensity(lngLineIndex).ToolTipText = Format(.XYDataList(StartXValIndex + XYDataIndex - 1).XVal, "0.0000")
                            ThisForm.lblPlotIntensity(lngLineIndex).Width = 50
                            ThisForm.lblPlotIntensity(lngLineIndex).Height = Abs(ThisForm.linData(lngLineIndex).y1 - ThisForm.linData(lngLineIndex).y2)
                            ThisForm.lblPlotIntensity(lngLineIndex).Top = ThisForm.linData(lngLineIndex).y2
                            ThisForm.lblPlotIntensity(lngLineIndex).Left = ThisForm.linData(lngLineIndex).x2
                        End If
                    Else
                        ThisForm.lblPlotIntensity(lngLineIndex).Visible = False
                    End If
                End With
            Next lngIndex
            
            intDynamicObjectOffset = intDynamicObjectOffset + XYDataToPlotCount
            
        End If
    Next intDataSetIndex
    
    With ThisForm
        If intDynamicObjectOffset < PlotOptions.LinesLoadedCount Then
            If intDynamicObjectOffset < 1 Then lngIndex = 1
            
            ' Hide the other lines and labels
            For lngLineIndex = intDynamicObjectOffset To PlotOptions.LinesLoadedCount
                ThisForm.linData(lngLineIndex).Visible = False
                ThisForm.lblPlotIntensity(lngLineIndex).Visible = False
            Next lngLineIndex
        End If
       
    End With
        
    If PlotOptions.ShowDataPointLabels Then
        TwipsBetweenLabels = "0.5" * DeltaXScaler * XValRange
        RepositionDataLabels ThisForm, XYDataToPlotCount, TwipsBetweenLabels, LastLabelTwip
    End If
    
End Sub

Private Sub FindMinimumAndMaximum(lngMinimumValIndex As Long, lngMaximumValIndex As Long, ThisXYData() As usrXYData, XYDataPointerArray() As Long, lngStartIndex As Long, lngStopIndex As Long)
    Dim lngIndex As Long, XYDataPoint As Long
    Dim dblMinimumVal As Double, dblMaximumVal As Double
    
    If lngStopIndex > UBound(ThisXYData()) Then
        lngStopIndex = UBound(ThisXYData())
    End If
    
    lngMinimumValIndex = lngStartIndex
    lngMaximumValIndex = lngStartIndex
    
    dblMaximumVal = ThisXYData(XYDataPointerArray(lngStartIndex)).YVal
    dblMinimumVal = ThisXYData(XYDataPointerArray(lngStartIndex)).YVal
    
    For lngIndex = lngStartIndex + 1 To lngStopIndex
        XYDataPoint = ThisXYData(XYDataPointerArray(lngIndex)).YVal
        If XYDataPoint < dblMinimumVal Then
            dblMinimumVal = XYDataPoint
            lngMinimumValIndex = lngIndex
        End If
        If XYDataPoint > dblMaximumVal Then
            dblMaximumVal = XYDataPoint
            lngMaximumValIndex = lngIndex
        End If
    Next lngIndex
    
    If lngMinimumValIndex = lngMaximumValIndex Then
        ' All of the data was the same
        ' Set lngMaximumValIndex to the halfway-point index between lngStopIndex and lngStartIndex
        lngMaximumValIndex = lngStartIndex + (lngStopIndex - lngStartIndex) / 2
    End If
End Sub

Private Sub FormatAxes(ThisForm As Form, PlotOptions As usrPlotDataOptions, PlotBottom As Long, PlotRange As usrPlotRange)
    Dim intInitialTickIndex As Integer, intInitialTickMinorIndex As Integer
    Dim lngRightOrTopMostPos As Long
    
    With ThisForm
        ' Position the x axis
        .linXAxis.x1 = PlotOptions.PlotLeft + PlotOptions.PlotLeftLargeNumberOffset
        .linXAxis.x2 = PlotOptions.PlotLeft + PlotOptions.PlotWidth + 50
        .linXAxis.y1 = PlotBottom + 50
        .linXAxis.y2 = PlotBottom + 50
        .linXAxis.Visible = PlotOptions.XAxis.Show
        
        ' Position the y axis
        .linYAxis.x1 = PlotOptions.PlotLeft + PlotOptions.PlotLeftLargeNumberOffset
        .linYAxis.x2 = PlotOptions.PlotLeft + PlotOptions.PlotLeftLargeNumberOffset
        .linYAxis.y1 = PlotBottom + 50
        .linYAxis.y2 = PlotOptions.PlotTop - 50
        .linYAxis.Visible = PlotOptions.YAxis.Show
    End With
    
    ' Note: The x and y axes share the same dynamic lines for major and minor tick marks,
    '       gridlines, and labels.  The x axis objects will start with index 1 of each object type
    '       the intInitialTickIndex and intInitialTickMinorIndex values will be modified by sub
    '       FormatThisAxis during the creation of the objects for the x axis so that they will
    '       be the value of the next unused object for operations involving the y axis
    intInitialTickIndex = 0
    intInitialTickMinorIndex = 0
    FormatThisAxis ThisForm, PlotOptions, True, PlotOptions.XAxis, PlotRange.x, intInitialTickIndex, intInitialTickMinorIndex, lngRightOrTopMostPos
    If lngRightOrTopMostPos > ThisForm.linXAxis.x2 Then
        ThisForm.linXAxis.x2 = lngRightOrTopMostPos + 50
    End If
    
    FormatThisAxis ThisForm, PlotOptions, False, PlotOptions.YAxis, PlotRange.y, intInitialTickIndex, intInitialTickMinorIndex, lngRightOrTopMostPos
    If lngRightOrTopMostPos < ThisForm.linYAxis.y2 Then
       ThisForm.linYAxis.y2 = lngRightOrTopMostPos - 50
    End If
   
End Sub
    
Private Function FindDigitsInLabelUsingRange(PlotRangeForAxis As usrPlotRangeAxis, intMajorTicksToShow As Integer, ByRef strFormatString As String) As Integer
    Dim ValRange As Double, DeltaVal As Double, lngExponentValue As Long
    Dim intDigitsInStartNumber As Integer, intDigitsInEndNumber As Integer
    
    With PlotRangeForAxis
        ValRange = .ValEnd.Val - .ValStart.Val
    End With
    
    If intMajorTicksToShow = 0 Then
        DeltaVal = ValRange
    Else
        ' Take ValRange divided by intMajorTicksToShow
        DeltaVal = ValRange / intMajorTicksToShow
    End If
    
    ' Round DeltaVal to nearest 1, 2, or 5 (or multiple of 10 thereof)
    DeltaVal = RoundToMultipleOf10(DeltaVal, lngExponentValue)
    
    strFormatString = ConstructFormatString(DeltaVal)
    
    ' Note: I use the absolute value of ValEnd so that all tick labels will have the
    ' same width, positive or negative
    intDigitsInStartNumber = Len(Trim(Format(Abs(PlotRangeForAxis.ValStart.Val), strFormatString))) + 1
    intDigitsInEndNumber = Len(Trim(Format(Abs(PlotRangeForAxis.ValEnd.Val), strFormatString))) + 1
    
    If intDigitsInEndNumber > intDigitsInStartNumber Then
        FindDigitsInLabelUsingRange = intDigitsInEndNumber
    Else
        FindDigitsInLabelUsingRange = intDigitsInStartNumber
    End If
    
End Function

Private Sub FormatThisAxis(ThisForm As Form, PlotOptions As usrPlotDataOptions, boolXAxis As Boolean, AxisOptions As usrAxisOptions, PlotRangeForAxis As usrPlotRangeAxis, ByRef intInitialTickIndex As Integer, ByRef intInitialTickMinorIndex As Integer, lngRightOrTopMostPos As Long)
    Dim objTickMajor As Line
    
    Dim intMajorTicksToShow As Integer
    Dim intMinorTicksPerMajorTick As Integer, intMinorTicksRequired As Integer
    Dim intTickIndex As Integer, intTickIndexToHide As Integer
    Dim intKeepEveryXthTick As Integer
    Dim intTickIndexMinor As Integer, intTickIndexMinorTrack As Integer
    Dim lngAddnlMinorTickPos As Long, lngAddnlMinorTickStopPos As Long
    Dim PosStart As Long, DeltaPos As Double, DeltaPosMinor As Double
    Dim ValStart As Double, ValEnd As Double, ValRange As Double, DeltaVal As Double
    Dim LengthStartPos As Long, LengthEndPos As Long, GridlineEndPos As Long
    Dim LengthStartPosMinor As Long, LengthEndPosMinor As Long
    Dim intDigitsInLabel As Integer
    
    Dim strFormatString As String, intTickLabelWidth As Integer
    
    ' Position and label the x axis major tick marks and labels
    intMajorTicksToShow = AxisOptions.MajorTicksToShow
    If intMajorTicksToShow < 2 Then
        intMajorTicksToShow = 2
        AxisOptions.MajorTicksToShow = intMajorTicksToShow
    End If
    
    With PlotRangeForAxis
        ValStart = .ValStart.Val
        ValEnd = .ValEnd.Val
        ValRange = ValEnd - ValStart
    End With
    
    ' Take ValRange divided by intMajorTicksToShow
    DeltaVal = ValRange / intMajorTicksToShow
    
    ' Round DeltaVal to nearest 1, 2, or 5 (or multiple of 10 thereof)
    DeltaVal = RoundToMultipleOf10(DeltaVal)
    
    ' If ValStart is not an even multiple of DeltaVal then bump up ValStart until it is
    ValStart = RoundToEvenMultiple(ValStart, DeltaVal, True)
    
    ' Do the same for ValEnd, but bump down instead
    ValEnd = RoundToEvenMultiple(ValEnd, DeltaVal, False)
    
    ' Recompute ValRange
    ValRange = ValEnd - ValStart
    
    ' Determine actual number of ticks to show
    intMajorTicksToShow = ValRange / DeltaVal + 1
    
    ' Convert ValStart to ValPos
    PosStart = XYValueToPos(ValStart, PlotRangeForAxis, False)
    
    ' Convert DeltaVal to DeltaPos
    DeltaPos = XYValueToPos(DeltaVal, PlotRangeForAxis, True)
    
    If boolXAxis Then
        LengthStartPos = ThisForm.linXAxis.y1 + 200
        LengthEndPos = ThisForm.linXAxis.y1
    Else
        LengthStartPos = ThisForm.linYAxis.x1 - 200
        LengthEndPos = ThisForm.linYAxis.x1
    End If
        
    With AxisOptions
        If .ShowMinorTicks Then
            ' Insert 1, 4, or 9 minor ticks, depending on whether they'll fit
            intMinorTicksPerMajorTick = 0
            Do
                Select Case intMinorTicksPerMajorTick
                Case 0: intMinorTicksPerMajorTick = 9
                Case 9: intMinorTicksPerMajorTick = 4
                Case 4: intMinorTicksPerMajorTick = 1
                Case Else
                    intMinorTicksPerMajorTick = 0
                    Exit Do
                End Select
                DeltaPosMinor = DeltaPos / (intMinorTicksPerMajorTick + 1)
            Loop While Abs(DeltaPosMinor) < .MinorTickMinimumPixelSep
        Else
            DeltaPosMinor = 0
        End If
        
        If boolXAxis Then
            LengthStartPosMinor = LengthStartPos - 150
            LengthEndPosMinor = LengthEndPos
        Else
            LengthStartPosMinor = LengthStartPos + 100
            LengthEndPosMinor = LengthEndPos
        End If
        
        If .ShowGridLinesMajor Then
            If boolXAxis Then
                GridlineEndPos = PlotOptions.PlotTop
            Else
                GridlineEndPos = PlotOptions.PlotLeft + PlotOptions.PlotWidth
            End If
        Else
            GridlineEndPos = LengthEndPos
        End If
    End With
    
    ' Initialize PlotOptions.MajorTicksLoaded and PlotOptions.MinorTicksLoaded if needed
    If PlotOptions.MajorTicksLoaded = 0 Then PlotOptions.MajorTicksLoaded = 1 ' There is always at least 1 loaded
    If PlotOptions.MinorTicksLoaded = 0 Then PlotOptions.MinorTicksLoaded = 1  ' There is always at least 1 loaded

    ' Call FindDigitsInLabelUsingRange to determine the digits in the label and construct the format string
    intDigitsInLabel = FindDigitsInLabelUsingRange(PlotRangeForAxis, intMajorTicksToShow, strFormatString)
    
    ' Each number requires 90 pixels
    intTickLabelWidth = intDigitsInLabel * NUM_TWIPS_PER_DIGIT

    intTickIndexMinorTrack = intInitialTickMinorIndex
    For intTickIndex = intInitialTickIndex + 1 To intInitialTickIndex + intMajorTicksToShow
    
        With AxisOptions
            If PlotOptions.MajorTicksLoaded < intTickIndex Then
                PlotOptions.MajorTicksLoaded = PlotOptions.MajorTicksLoaded + 1
                Load ThisForm.linTickMajor(PlotOptions.MajorTicksLoaded)
                Load ThisForm.lblTick(PlotOptions.MajorTicksLoaded)
                Load ThisForm.linGridline(PlotOptions.MajorTicksLoaded)
            End If
        End With
        
        With ThisForm.linTickMajor(intTickIndex)
            .x1 = PosStart + CLng(((intTickIndex - intInitialTickIndex) - 1) * DeltaPos)
            .x2 = .x1
            .y1 = LengthStartPos
            .y2 = LengthEndPos
            .Visible = True
        End With
        Set objTickMajor = ThisForm.linTickMajor(intTickIndex)
        With ThisForm.linGridline(intTickIndex)
            .x1 = objTickMajor.x1
            .x2 = objTickMajor.x2
            .y1 = objTickMajor.y1
            .y2 = GridlineEndPos
            .Visible = True
        End With
        
        lngRightOrTopMostPos = LengthEndPos
        If Not boolXAxis Then
            SwapLineCoordinates ThisForm.linTickMajor(intTickIndex)
            SwapLineCoordinates ThisForm.linGridline(intTickIndex)
            Set objTickMajor = ThisForm.linTickMajor(intTickIndex)
        End If
        
        With ThisForm.lblTick(intTickIndex)
            .Width = intTickLabelWidth
            .Caption = Format(ValStart + CDbl((intTickIndex - intInitialTickIndex) - 1) * DeltaVal, strFormatString)
            .Visible = True
            If boolXAxis Then
                .Top = objTickMajor.y1 + 50
                .Left = objTickMajor.x1 - .Width / 2
            Else
                .Top = objTickMajor.y1 - .Height / 2
                .Left = objTickMajor.x1 - .Width - 50
                .Alignment = vbRightJustify
            End If
        End With
        
        If ValStart + ((intTickIndex - intInitialTickIndex) - 1) * DeltaVal > PlotRangeForAxis.ValEnd.Val Then
            ' Tick and/or label is past the end of the axis; Do not keep labeling
            ' Before exitting for loop, must add 1 to intTickIndex since it will normally
            '  exit the for loop with a value of intMajorTicksToShow + 1, but we are exiting prematurely
            intTickIndex = intTickIndex + 1
            Exit For
        End If

        With AxisOptions
            ' Load minor ticks as needed
            intMinorTicksRequired = intTickIndexMinorTrack + intMinorTicksPerMajorTick
            Do While PlotOptions.MinorTicksLoaded < intMinorTicksRequired
                PlotOptions.MinorTicksLoaded = PlotOptions.MinorTicksLoaded + 1
                Load ThisForm.linTickMinor(PlotOptions.MinorTicksLoaded)
            Loop
        End With
        
        For intTickIndexMinor = 1 To intMinorTicksPerMajorTick
            intTickIndexMinorTrack = intTickIndexMinorTrack + 1
            If boolXAxis Then
                lngAddnlMinorTickPos = objTickMajor.x1 + DeltaPosMinor * intTickIndexMinor
            Else
                lngAddnlMinorTickPos = objTickMajor.y1 + DeltaPosMinor * intTickIndexMinor
            End If
            
            lngRightOrTopMostPos = lngAddnlMinorTickPos
            
            ' Add the minor tick mark
            AddMinorTickmark ThisForm, intTickIndexMinor, intTickIndexMinorTrack, lngAddnlMinorTickPos, LengthStartPosMinor, LengthEndPosMinor, boolXAxis
        Next intTickIndexMinor
    Next intTickIndex
    
    ' See if minor ticks can be added before the first major tick
    If AxisOptions.ShowMinorTicks Then
        If boolXAxis Then
            lngAddnlMinorTickPos = ThisForm.linTickMajor(1).x1
            lngAddnlMinorTickStopPos = ThisForm.linXAxis.x1
        Else
            lngAddnlMinorTickPos = ThisForm.linTickMajor(intInitialTickIndex + 1).y1
            lngAddnlMinorTickStopPos = ThisForm.linYAxis.y1
        End If
        
        intTickIndexMinor = 1
        ' This loop will execute at most 50 times
        ' Most likely, it will only execute a few times before the Exit Do clause becomes true
        Do While intTickIndexMinor < 50
            If boolXAxis Then
                If lngAddnlMinorTickPos - Abs(DeltaPosMinor) <= lngAddnlMinorTickStopPos Then Exit Do
                lngAddnlMinorTickPos = lngAddnlMinorTickPos - Abs(DeltaPosMinor)
            Else
                If lngAddnlMinorTickPos + Abs(DeltaPosMinor) >= lngAddnlMinorTickStopPos Then Exit Do
                lngAddnlMinorTickPos = lngAddnlMinorTickPos + Abs(DeltaPosMinor)
            End If
            
            intTickIndexMinor = intTickIndexMinor + 1
            intTickIndexMinorTrack = intTickIndexMinorTrack + 1
            
            ' May need to load more minor ticks
            If PlotOptions.MinorTicksLoaded < intTickIndexMinorTrack Then
                PlotOptions.MinorTicksLoaded = PlotOptions.MinorTicksLoaded + 1
                Load ThisForm.linTickMinor(PlotOptions.MinorTicksLoaded)
            End If

            ' Add the minor tick mark
            AddMinorTickmark ThisForm, intTickIndexMinor, intTickIndexMinorTrack, lngAddnlMinorTickPos, LengthStartPosMinor, LengthEndPosMinor, boolXAxis
        Loop
    End If
    
    With ThisForm
        ' Check for overlapping tick labels
        ' If there is overlap, then show every other or every 5th label
        intKeepEveryXthTick = 1
        If intMajorTicksToShow >= 3 Then
            If boolXAxis Then
                If .lblTick(1).Left + .lblTick(1).Width > .lblTick(3).Left Then intKeepEveryXthTick = 5
            Else
                If .lblTick(intInitialTickIndex + 3).Top + .lblTick(intInitialTickIndex + 3).Height > .lblTick(intInitialTickIndex + 1).Top Then intKeepEveryXthTick = 5
            End If
        End If
        
        If intKeepEveryXthTick = 1 And intMajorTicksToShow >= 2 Then
            If boolXAxis Then
                If .lblTick(1).Left + .lblTick(1).Width > .lblTick(2).Left Then intKeepEveryXthTick = 2
            Else
                If .lblTick(intInitialTickIndex + 2).Top + .lblTick(intInitialTickIndex + 2).Height > .lblTick(intInitialTickIndex + 1).Top Then intKeepEveryXthTick = 2
            End If
        End If
        
        If intKeepEveryXthTick > 1 Then
            For intTickIndexToHide = intInitialTickIndex + 1 To intInitialTickIndex + intMajorTicksToShow
                If intTickIndexToHide Mod intKeepEveryXthTick <> 0 Then
                    .lblTick(intTickIndexToHide).Visible = False
                End If
            Next intTickIndexToHide
        End If
    End With
    
    ' Store intTickIndex in intInitialTickIndex
    intInitialTickIndex = intTickIndex - 1
    
    ' Store intTickIndexMinorTrack in intInitialTickMinorIndex
    intInitialTickMinorIndex = intTickIndexMinorTrack
    
    
    With ThisForm
        ' Hide remaining loaded ticks
        For intTickIndexToHide = intInitialTickIndex + 1 To PlotOptions.MajorTicksLoaded
            .linTickMajor(intTickIndexToHide).Visible = False
            .linGridline(intTickIndexToHide).Visible = False
            .lblTick(intTickIndexToHide).Visible = False
        Next intTickIndexToHide
        
        For intTickIndexToHide = intInitialTickMinorIndex + 1 To PlotOptions.MinorTicksLoaded
            .linTickMinor(intTickIndexToHide).Visible = False
        Next intTickIndexToHide
    End With
    
End Sub

Private Sub AddMinorTickmark(ThisForm As Form, intTickIndexMinor As Integer, intTickIndexMinorTrack As Integer, lngAddnlMinorTickPos As Long, LengthStartPosMinor As Long, LengthEndPosMinor As Long, boolXAxis As Boolean)
    With ThisForm.linTickMinor(intTickIndexMinorTrack)
        .x1 = lngAddnlMinorTickPos
        .x2 = .x1
        .y1 = LengthStartPosMinor + 20
        .y2 = LengthEndPosMinor
        If intTickIndexMinor Mod 5 = 0 Then
            ' Draw the minor tick mark a little longer
            If boolXAxis Then
                .y1 = .y1 + 70
                .y2 = .y2
            Else
                .y1 = .y1 - 50
                .y2 = .y2
            End If
        End If
        .Visible = True
    End With

    If Not boolXAxis Then SwapLineCoordinates ThisForm.linTickMinor(intTickIndexMinorTrack)

End Sub

Private Sub LoadDynamicPlotObjects(ThisForm As Form, PlotOptions As usrPlotDataOptions, XYDataToPlotCount As Long)
    Dim lngLinesLoadedCountPrevious  As Long
    Dim boolShowProgress As Boolean
    
    With ThisForm
        ' Load lines for data points if needed
        If PlotOptions.LinesLoadedCount < XYDataToPlotCount Then
            lngLinesLoadedCountPrevious = PlotOptions.LinesLoadedCount
            If Abs(XYDataToPlotCount - lngLinesLoadedCountPrevious) > 100 Then
                boolShowProgress = True
                frmProgress.InitializeForm "Preparing graph", 0, Abs(XYDataToPlotCount - lngLinesLoadedCountPrevious)
            End If
            
            Do While PlotOptions.LinesLoadedCount < XYDataToPlotCount
                If boolShowProgress Then
                    If PlotOptions.LinesLoadedCount Mod 10 = 0 Then frmProgress.UpdateProgressBar Abs(PlotOptions.LinesLoadedCount - lngLinesLoadedCountPrevious)
                End If
                
                PlotOptions.LinesLoadedCount = PlotOptions.LinesLoadedCount + 1
                Load .linData(PlotOptions.LinesLoadedCount)
                Load .lblPlotIntensity(PlotOptions.LinesLoadedCount)
                .lblPlotIntensity(PlotOptions.LinesLoadedCount).Height = 200
            Loop
            If boolShowProgress Then frmProgress.Hide
        End If
    End With
    
End Sub

Private Sub RepositionDataLabels(ThisForm As Form, XYDataToPlotCount As Long, TwipsBetweenLabels As Long, LastLabelTwip As Long)
    Dim LabelLocPointerArray() As Long, LabelLocPointerArrayCount As Long
    Dim LabelShiftedCount As Integer, LabelShifted As Boolean
    
    Dim lngIndex As Integer
    
    With ThisForm
        ' Reposition the labels
        ReDim LabelLocPointerArray(XYDataToPlotCount)
        LabelLocPointerArrayCount = 0
        For lngIndex = 1 To XYDataToPlotCount
            If .lblPlotIntensity(lngIndex).Tag = "Visible" Then
                LabelLocPointerArrayCount = LabelLocPointerArrayCount + 1
                LabelLocPointerArray(LabelLocPointerArrayCount) = lngIndex
            End If
        Next lngIndex
        
        LabelShiftedCount = 0
        Do
            LabelShiftedCount = LabelShiftedCount + 1
            LabelShifted = False
            ' sort the Pointer Array
            SortLabelLoc ThisForm, LabelLocPointerArray(), LabelLocPointerArrayCount
            
            ' Step through labels from the one at the bottom to the one at the top and shift upward if needed
            For lngIndex = 1 To LabelLocPointerArrayCount - 1
                If .lblPlotIntensity(LabelLocPointerArray(lngIndex + 1)).Top + .lblPlotIntensity(LabelLocPointerArray(lngIndex + 1)).Height > .lblPlotIntensity(lngIndex).Top And .lblPlotIntensity(lngIndex + 1).Top < .lblPlotIntensity(LabelLocPointerArray(lngIndex)).Top Then
                    ' May Need to shift upward; see if adjacent
                    If Abs(LabelLocPointerArray(lngIndex + 1) - LabelLocPointerArray(lngIndex)) = 1 Then
                        ' Yes, they're adjacent; shift upward
                        .lblPlotIntensity(LabelLocPointerArray(lngIndex + 1)).Top = .lblPlotIntensity(LabelLocPointerArray(lngIndex)).Top - .lblPlotIntensity(LabelLocPointerArray(lngIndex + 1)).Height
                        LabelShifted = True
                    End If
                End If
            Next lngIndex
        Loop While LabelShifted And LabelShiftedCount < 5
    End With
    
End Sub

Private Function RoundToEvenMultiple(ByVal dblValueToRound As Double, ByVal MultipleValue As Double, ByVal boolRoundUp As Boolean) As Double
    Dim intLoopCount As Integer
    Dim strWork As String, dblWork As Double
    Dim lngExponentValue As Long
    
    ' Find the exponent of MultipleValue
    strWork = Format(MultipleValue, "0E+000")
    lngExponentValue = CIntSafe(Right(strWork, 4))
    
    intLoopCount = 0
    Do While Trim(Str(dblValueToRound / MultipleValue)) <> Trim(Str(Round(dblValueToRound / MultipleValue, 0)))
        dblWork = dblValueToRound / 10 ^ (lngExponentValue)
        dblWork = CLng(dblWork)
        dblWork = dblWork * 10 ^ (lngExponentValue)
        If boolRoundUp Then
            If dblWork <= dblValueToRound Then
                dblWork = dblWork + 10 ^ lngExponentValue
            End If
        Else
            If dblWork >= dblValueToRound Then
                dblWork = dblWork - 10 ^ lngExponentValue
            End If
        End If
        dblValueToRound = dblWork
        intLoopCount = intLoopCount + 1
        If intLoopCount > 500 Then
            ' Bug
'            Debug.Assert False
            Exit Do
        End If
    Loop
    
    RoundToEvenMultiple = dblValueToRound
End Function

Public Function RoundToMultipleOf10(ByVal dblThisNum As Double, Optional ByRef lngExponentValue As Long) As Double
    Dim strWork As String, dblWork As Double
    
    ' Round to nearest 1, 2, or 5 (or multiple of 10 thereof)
    ' First, find the exponent of dblThisNum
    strWork = Format(dblThisNum, "0E+000")
    lngExponentValue = CIntSafe(Right(strWork, 4))
    dblWork = dblThisNum / 10 ^ lngExponentValue
    dblWork = CIntSafeDbl(dblWork)
    
    ' dblWork should now be between 0 and 9
    Select Case dblWork
    Case 0, 1: dblThisNum = 1
    Case 2 To 4: dblThisNum = 2
    Case Else: dblThisNum = 5
    End Select
    
    ' Convert dblThisNum back to the correct magnitude
    dblThisNum = dblThisNum * 10 ^ lngExponentValue
    
    RoundToMultipleOf10 = dblThisNum
End Function

Private Sub ScaleData(PlotOptions As usrPlotDataOptions, XYDataToPlot() As usrXYData, XYDataToPlotCount As Long, ByRef ThisAxisRange As usrPlotRangeAxis, ThisAxisScaling As usrPlotRangeAxis, boolIsXAxis As Boolean, boolAutoScaleAxis As Boolean)

    Dim dblMinimumIntensity As Double
    Dim lngIndex As Long, dblDataPoint As Double
    Dim dblValRange As Double, dblDeltaScaler As Double, ThisAxisWindowLength As Long
    Dim strFormatString As String, intDigitsInLabel As Integer
    
    ' First step through valid data and find the minimum YVal value
    ' Necessary to check if negative
    dblMinimumIntensity = HighestValueForDoubleDataType
    For lngIndex = 1 To XYDataToPlotCount
        If boolIsXAxis Then
            dblDataPoint = XYDataToPlot(lngIndex).XVal
        Else
            dblDataPoint = XYDataToPlot(lngIndex).YVal
        End If
        If dblDataPoint < dblMinimumIntensity Then
            dblMinimumIntensity = dblDataPoint
        End If
    Next lngIndex
    
    ' Reset .ValNegativeValueCorrectionOffset
    ThisAxisRange.ValNegativeValueCorrectionOffset = 0
    
    If boolAutoScaleAxis Then
        If dblMinimumIntensity < 0 Then
            ' Need to correct all y data by making it positive
            ThisAxisRange.ValNegativeValueCorrectionOffset = Abs(dblMinimumIntensity)
        End If
    Else
        ' The user has supplied ValStart and ValEnd values
        ' Make sure .ValStart.Val < .ValEnd.Val
        ' No need to use a NegativeValueCorrectionOffset since I perform bounds checking during the conversion from val to pos
        If ThisAxisScaling.ValStart.Val > ThisAxisScaling.ValEnd.Val Then
            SwapValues ThisAxisScaling.ValStart.Val, ThisAxisScaling.ValEnd.Val
        End If
    End If
    
    If ThisAxisRange.ValNegativeValueCorrectionOffset > 0 Then
        For lngIndex = 1 To XYDataToPlotCount
            If boolIsXAxis Then
                XYDataToPlot(lngIndex).XVal = XYDataToPlot(lngIndex).XVal + ThisAxisRange.ValNegativeValueCorrectionOffset
            Else
                XYDataToPlot(lngIndex).YVal = XYDataToPlot(lngIndex).YVal + ThisAxisRange.ValNegativeValueCorrectionOffset
            End If
        Next lngIndex
    End If
    
    ' Record the current plot range for future reference when zooming
    ThisAxisRange.ValStart.Val = ThisAxisScaling.ValStart.Val - ThisAxisRange.ValNegativeValueCorrectionOffset
    ThisAxisRange.ValEnd.Val = ThisAxisScaling.ValEnd.Val - ThisAxisRange.ValNegativeValueCorrectionOffset
    
    dblValRange = ThisAxisScaling.ValEnd.Val - ThisAxisScaling.ValStart.Val
    If dblValRange = 0 Then dblValRange = 1
    
    ' Scale the data according to PlotOptions.height
    
    If boolIsXAxis Then
        ThisAxisWindowLength = PlotOptions.PlotWidth
    Else
        ThisAxisWindowLength = PlotOptions.PlotHeight
    End If
        
    dblDeltaScaler = CDbl(ThisAxisWindowLength) / CDbl(dblValRange)
    
    For lngIndex = 1 To XYDataToPlotCount
        If boolIsXAxis Then
            dblDataPoint = XYDataToPlot(lngIndex).XVal
        Else
            dblDataPoint = XYDataToPlot(lngIndex).YVal
        End If
        
        dblDataPoint = CLng((dblDataPoint - ThisAxisScaling.ValStart.Val) * dblDeltaScaler)
        If dblDataPoint > ThisAxisWindowLength Then
            dblDataPoint = ThisAxisWindowLength
        Else
            If dblDataPoint < 0 Then
                dblDataPoint = 0
            End If
        End If
        
        If boolIsXAxis Then
            XYDataToPlot(lngIndex).XVal = dblDataPoint
        Else
            XYDataToPlot(lngIndex).YVal = dblDataPoint
        End If

    Next lngIndex
    
    If Not boolIsXAxis Then
        ' Need to recompute .PlotLeftLargeNumberOffset
        ' Call FindDigitsInLabelUsingRange to determine the digits in the label and construct the format string
        intDigitsInLabel = FindDigitsInLabelUsingRange(ThisAxisRange, PlotOptions.YAxis.MajorTicksToShow, strFormatString)
        
        ' Store in .PlotLeftLargeNumberOffset
        PlotOptions.PlotLeftLargeNumberOffset = intDigitsInLabel * NUM_TWIPS_PER_DIGIT
    End If
    
    ' Record the current plot range for future reference when zooming
    If boolIsXAxis Then
        ThisAxisRange.ValStart.Pos = PlotOptions.PlotLeft + PlotOptions.PlotLeftLargeNumberOffset
        ThisAxisRange.ValEnd.Pos = PlotOptions.PlotLeft + PlotOptions.PlotLeftLargeNumberOffset + PlotOptions.PlotWidth
    Else
        ThisAxisRange.ValStart.Pos = PlotOptions.PlotTop + PlotOptions.PlotHeight
        ThisAxisRange.ValEnd.Pos = PlotOptions.PlotTop     ' equivalent to: PlotBottom - CLng((dblValRange - ThisAxisScaling.ValStart) * dblDeltaScaler)
    End If
    
End Sub

Private Sub SortLabelLoc(ThisForm As Form, LabelLocPointerArray() As Long, LabelLocPointerArrayCount As Long)
    
    ' Sorts a list of labels based on their location vertically on a form
    Dim low%, high%, IndexTemp As Long
    
    low% = 1
    high% = LabelLocPointerArrayCount
    
    ' Sort the list via a shell sort
    Dim MaxRow As Integer, offset As Integer, limit As Integer, switch As Integer
    Dim row As Integer
    
    ' Set comparison offset to half the number of records
    MaxRow = high
    offset = MaxRow \ 2

    Do While offset > 0          ' Loop until offset gets to zero.

        limit = MaxRow - offset
        Do
            switch = 0         ' Assume no switches at this offset.

            ' Compare elements and switch ones out of order:
            For row = low To limit
                If ThisForm.lblPlotIntensity(LabelLocPointerArray(row)).Top < ThisForm.lblPlotIntensity(LabelLocPointerArray(row + offset)).Top Then
                    IndexTemp = LabelLocPointerArray(row + offset)
                    LabelLocPointerArray(row + offset) = LabelLocPointerArray(row)
                    LabelLocPointerArray(row) = IndexTemp
                    switch = row
                End If
            Next row

            ' Sort on next pass only to where last switch was made:
            limit = switch - offset
        Loop While switch

        ' No switches at last offset, try one half as big:
        offset = offset \ 2
    
    Loop

End Sub

Private Sub SortXYData(XYData() As usrXYData, XYDataPointerArray() As Long, XYDataPointerCount As Long, boolSortByIntensity As Boolean)
        
    ' Sorts a list of data by x value or by YVal (y values), depending on boolSortByIntensity
    
    ' Rather than sorting the data itself, sorts a pointer array to the data
    Dim lngLow As Long, lngHigh As Long, IndexTemp As Long, boolSwapThem As Boolean
    
    lngLow = 1
    lngHigh = XYDataPointerCount
    
    ' Sort the list via a shell sort
    Dim lngRowMax As Long, lngOffSet As Long, lngLimit As Long, lngSwitch As Long
    Dim lngRow As Long
    
    ' Set comparison lngOffSet to half the number of records
    lngRowMax = lngHigh
    lngOffSet = lngRowMax \ 2

    Do While lngOffSet > 0          ' Loop until lngOffSet gets to zero.

        lngLimit = lngRowMax - lngOffSet
        Do
            lngSwitch = 0         ' Assume no switches at this lngOffSet.

            ' Compare elements and lngSwitch ones out of order:
            For lngRow = lngLow To lngLimit
                If boolSortByIntensity Then
                    boolSwapThem = (XYData(XYDataPointerArray(lngRow)).YVal < XYData(XYDataPointerArray(lngRow + lngOffSet)).YVal)
                Else
                    boolSwapThem = (XYData(XYDataPointerArray(lngRow)).XVal > XYData(XYDataPointerArray(lngRow + lngOffSet)).XVal)
                End If
                
                If boolSwapThem Then
                    IndexTemp = XYDataPointerArray(lngRow + lngOffSet)
                    XYDataPointerArray(lngRow + lngOffSet) = XYDataPointerArray(lngRow)
                    XYDataPointerArray(lngRow) = IndexTemp
                    lngSwitch = lngRow
                End If
            Next lngRow

            ' Sort on next pass only to where last lngSwitch was made:
            lngLimit = lngSwitch - lngOffSet
        Loop While lngSwitch

        ' No switches at last lngOffSet, try one half as big:
        lngOffSet = lngOffSet \ 2
    
    Loop

End Sub

Private Sub SwapLineCoordinates(ThisLine As Line)
    Dim lngTemp As Long
    
    With ThisLine
        ' Setting y axis values
        ' Must swap x1 and y1, and x2 and y2
        lngTemp = .x1
        .x1 = .y1
        .y1 = lngTemp
        
        lngTemp = .x2
        .x2 = .y2
        .y2 = lngTemp

    End With
End Sub
Public Function XYPosToValue(ThisPos As Long, ThisRange As usrPlotRangeAxis)
    Dim PosRange As Long, ValRange As Double
    Dim ScaledPos As Double
    
    With ThisRange
        ' Convert the x pos in ZoomBoxCoords to the actual x value
        PosRange = .ValEnd.Pos - .ValStart.Pos
        ValRange = .ValEnd.Val - .ValStart.Val
        
        If PosRange <> 0 Then
            ' ScaledPos is a value between 0 and 1 indicating the percentage between RangeEnd and RangeStart that ThisPos is
            ScaledPos = (ThisPos - .ValStart.Pos) / PosRange
            
            ' Now Convert to the displayed value
            XYPosToValue = ScaledPos * ValRange + .ValStart.Val
        Else
            XYPosToValue = .ValStart.Val
        End If
    End With
    
End Function

Public Function XYValueToPos(ThisValue As Double, ThisRange As usrPlotRangeAxis, IsDeltaValue As Boolean)
    Dim PosRange As Long, ValRange As Double
    Dim ScaledPos As Double
    
    
    With ThisRange
        ' Convert the x pos in ZoomBoxCoords to the actual x value
        PosRange = .ValEnd.Pos - .ValStart.Pos
        ValRange = .ValEnd.Val - .ValStart.Val
    
        If ValRange <> 0 Then
            If IsDeltaValue Then
                ScaledPos = ThisValue / ValRange
            
                ' Now Convert to correct position
                XYValueToPos = ScaledPos * PosRange
            Else
                ' ScaledPos is a value between 0 and 1 indicating the percentage between RangeEnd and RangeStart that ThisPos is
                ScaledPos = (ThisValue - .ValStart.Val) / ValRange
                
                ' Now Convert to correct position
                XYValueToPos = ScaledPos * PosRange + .ValStart.Pos
            End If
        Else
            XYValueToPos = ThisRange.ValStart.Pos
        End If
    End With
    
End Function


