Attribute VB_Name = "MsPlotRoutines"
Option Explicit

Private Const NUM_TWIPS_PER_DIGIT = 90
Public Const MAX_DATA_SETS = 2              ' Up to 2 data sets may be graphed simultaneously (uses indices 0 and 1)

Public Const cPlotTypeSticks = 0
Public Const cPlotTypeGaussian = 1

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
    Y As usrPlotRangeAxis
End Type

Public Type usrGaussianOptions
    ResolvingPower As Long              ' Effective resolution  (M / delta M)
    XValueOfSpecification As Double     ' X Value where effective resolution applies
    QualityFactor As Integer            ' The higher this value is, the more data points are created for each Gaussian peak
End Type

Public Type usrPlotDataOptions
    
    PlotTypeCode As Integer     ' Whether the plot is a stick plot (0) or line-between-points plot (1)
    GaussianConversion As usrGaussianOptions
    ApproximationFactor As Integer      ' Affects when and how many data points are discarded when approximating the graph
    
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

    blnLongOperationsRequired As Boolean       ' Set to true when long operations are encountered, thus requiring an hourglass cursor
    
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

Private Sub CheckForLongOperation(ThisForm As VB.Form, lngSecondsElapsedAtStart As Long, ThesePlotOptions As usrPlotDataOptions, lngCurrentIteration As Long, lngTotalIterations As Long, strCurrentTask As String)
    ' Checks to see if the current value of Timer() is greater than lngSecondsElapsedAtStart + 1
    '  If it is then blnLongOperationsRequired is turned on and the pointer is changed to an hourglass
    ' The blnLongOperationsRequired value is saved for future calls to the sub so that the hourglass
    '  will be activated immediately on future calls
    ' Furthermore, if over 2 seconds have elapsed, then a progress box is shown
    Dim lngSecElapsedSinceOperationStart As Long
    
    lngSecElapsedSinceOperationStart = Timer() - lngSecondsElapsedAtStart
    
    If lngSecElapsedSinceOperationStart >= 1 Or ThesePlotOptions.blnLongOperationsRequired Then
        ThesePlotOptions.blnLongOperationsRequired = True
        If ThisForm.MousePointer <> vbHourglass Then
            ThisForm.MousePointer = vbHourglass
        End If
        
        If lngSecElapsedSinceOperationStart >= 2 Then
            ' Process has taken over 2 seconds
            ' Show the progress form if not shown yet
            If frmProgress.Visible = False Then
                frmProgress.InitializeForm strCurrentTask, 0, lngTotalIterations
                frmProgress.ToggleAlwaysOnTop True
            End If
            frmProgress.UpdateProgressBar lngCurrentIteration + 1
        End If
    End If
    
End Sub

Private Sub CheckDynamicXYData(ByRef ThisXYDataList() As usrXYData, ThisXYDataListCount As Long, ByRef ThisXYDataListCountDimmed As Long, Optional lngIncrement As Long = 100)
    If ThisXYDataListCount > ThisXYDataListCountDimmed Then
        ThisXYDataListCountDimmed = ThisXYDataListCountDimmed + lngIncrement
        If ThisXYDataListCountDimmed < ThisXYDataListCount Then
            ThisXYDataListCountDimmed = ThisXYDataListCount
        End If
        ReDim Preserve ThisXYDataList(ThisXYDataListCountDimmed)
    End If
End Sub

Public Function ConvertStickDataToGaussian(ThisForm As VB.Form, ThisXYDataSet As usrXYDataSet, ThesePlotOptions As usrPlotDataOptions, intDataSetIndex As Integer) As usrXYDataSet
    ' Note: Assumes ThisXYDataSet is sorted in the x direction
    
    Dim lngDataIndex As Long, lngMidPointIndex As Long
    Dim lngStickIndex As Long, DeltaX As Double
    Dim intQualityFactor As Integer
    
    Dim dblXValRange As Double, dblXValWindowRange As Double, dblRangeWork As Double
    Dim dblMinimalXValOfWindow As Double, dblMinimalXValSpacing As Double
    
    Dim dblXOffSet As Double, sigma As Double
    Dim lngExponentValue As Long
    Dim lngSecondsElapsedAtStart As Long, strCurrentTask As String
    
    Dim XYSummation() As usrXYData, XYSummationCount As Long, XYSummationCountDimmed As Long
    Dim lngSummationIndex As Long, lngMinimalSummationIndex As Long
    Dim DataToAdd() As usrXYData, lngDataToAddCount As Long, blnAppendNewData As Boolean
    Dim ThisDataPoint As usrXYData
    
    If ThisXYDataSet.XYDataListCount = 0 Then Exit Function
    
    lngSecondsElapsedAtStart = Timer()
    
    If ThesePlotOptions.GaussianConversion.ResolvingPower < 1 Then
        ThesePlotOptions.GaussianConversion.ResolvingPower = 1
    End If
    
    XYSummationCount = 0
    XYSummationCountDimmed = 100
    ReDim XYSummation(XYSummationCountDimmed)
    
    With ThisXYDataSet
        ' Make sure the Y Data range is defined
        CheckYDataRange .XYDataList(), .XYDataListCount, ThesePlotOptions, intDataSetIndex
    
        ThesePlotOptions.DataLimits(intDataSetIndex).x.ValStart.Val = .XYDataList(1).XVal
        ThesePlotOptions.DataLimits(intDataSetIndex).x.ValEnd.Val = .XYDataList(.XYDataListCount).XVal
    End With
    
    With ThesePlotOptions
        dblXValRange = .DataLimits(intDataSetIndex).x.ValEnd.Val - .DataLimits(intDataSetIndex).x.ValStart.Val
        
        intQualityFactor = .GaussianConversion.QualityFactor
        
        If intQualityFactor < 1 Or intQualityFactor > 50 Then
            intQualityFactor = 20
        End If
        
        ' Set DeltaX using .ResolvingPower and .XValueOfSpecification
        ' Do not allow the DeltaX to be so small that the total points required > 100,000
        DeltaX = .GaussianConversion.XValueOfSpecification / .GaussianConversion.ResolvingPower / intQualityFactor
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
        strCurrentTask = LookupMessage(1130)
        
        For lngStickIndex = 1 To .XYDataListCount
            If lngStickIndex Mod 25 = 0 Then
                CheckForLongOperation ThisForm, lngSecondsElapsedAtStart, ThesePlotOptions, lngStickIndex, .XYDataListCount, strCurrentTask
                If KeyPressAbortProcess Then Exit For
            End If
            
            ' Search through XYSummation to determine the index of the smallest XValue with which
            '   data in DataToAdd could be combined
            lngMinimalSummationIndex = 1
            dblMinimalXValOfWindow = .XYDataList(lngStickIndex).XVal - (lngMidPointIndex - 1) * DeltaX
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
            ThisDataPoint = .XYDataList(lngStickIndex)
            
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
                blnAppendNewData = True
            Else
                blnAppendNewData = False
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
                            blnAppendNewData = True
                        End If
                        Exit For
                    End If
                Next lngSummationIndex
            End If
            
            If blnAppendNewData = True Then
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
            
            ' Choose the appropriate new .XVal
            dblRangeWork = XYSummation(lngSummationIndex + 1).XVal - XYSummation(lngSummationIndex).XVal
            If dblRangeWork < dblMinimalXValSpacing * 2 Then
                dblRangeWork = dblRangeWork / 2
            Else
                dblRangeWork = dblMinimalXValSpacing
            End If
            XYSummation(lngSummationIndex + 1).XVal = XYSummation(lngSummationIndex).XVal + dblRangeWork
            
            ' The new .YVal is the average of that at lngSummationIndex and that at lngSummationIndex + 1
            XYSummation(lngSummationIndex + 1).YVal = (XYSummation(lngSummationIndex).YVal + XYSummation(lngSummationIndex + 1).YVal) / 2
        End If
        lngSummationIndex = lngSummationIndex + 1
    Loop
    
    frmProgress.HideForm

    ' Reset the blnLongOperationsRequired bit
    ThesePlotOptions.blnLongOperationsRequired = False

    ' ReDim XYSummation to XYSummationCount since DrawPlot assumes this is the case
    XYSummationCountDimmed = XYSummationCount
    ReDim Preserve XYSummation(XYSummationCountDimmed)
    
    ' Assign data in XYSummation to the function so that it gets returned
    ConvertStickDataToGaussian.XYDataList = XYSummation
    ConvertStickDataToGaussian.XYDataListCount = XYSummationCount
    ConvertStickDataToGaussian.XYDataListCountDimmed = XYSummationCountDimmed

End Function

Private Sub CheckYDataRange(ThisXYData() As usrXYData, ThisXYDataCount As Long, ThesePlotOptions As usrPlotDataOptions, intDataSetIndex As Integer)
    Dim lngIndex As Long
    Dim dblMaximumIntensity As Double, dblMinimumIntensity As Double
    Dim dblXYDataPoint As Double
    
    If ThesePlotOptions.DataLimits(intDataSetIndex).Y.ValStart.Val = 0 And ThesePlotOptions.DataLimits(intDataSetIndex).Y.ValEnd.Val = 0 Then
        ' Data Limits not defined
        ' Figure out what they are and sort data if necessary
        
        ' Find the Y scale data limits
        ' At the same time, see if the data is sorted
        dblMaximumIntensity = LowestValueForDoubleDataType
        dblMinimumIntensity = HighestValueForDoubleDataType
        For lngIndex = 1 To ThisXYDataCount
            dblXYDataPoint = ThisXYData(lngIndex).YVal
            If dblXYDataPoint > dblMaximumIntensity Then dblMaximumIntensity = dblXYDataPoint
            If dblXYDataPoint < dblMinimumIntensity Then dblMinimumIntensity = dblXYDataPoint
        Next lngIndex
        ThesePlotOptions.DataLimits(intDataSetIndex).Y.ValStart.Val = dblMinimumIntensity
        ThesePlotOptions.DataLimits(intDataSetIndex).Y.ValEnd.Val = dblMaximumIntensity
    End If

End Sub

Public Sub DrawPlot(ThisForm As VB.Form, ByRef ThesePlotOptions As usrPlotDataOptions, ByRef ThisXYDataArray() As usrXYDataSet, ByRef PlotRange() As usrPlotRange, intDataSetsLoaded As Integer)
    ' Draw a graphical representation of a list of x,y data pairs in one or more data sets (stored in ThisXYDataArray)
    ' Assumes the x,y data point array is 1-based (i.e. the first data point is in index 1
    ' Note: Assumes the data is sorted in the x direction
    
    ' ThesePlotOptions.Scaling is the axis range the user wishes to see
    ' The range actually used to display the data is stored in PlotRange
    
    Dim intDataSetIndex As Integer
    Dim lngIndex As Long, lngLineIndex As Long
    Dim PlotBottom As Long, lngLeftOffset As Long
    Dim lngSecondsElapsedAtStart As Long
    Dim intKeepEveryXthPoint As Integer, lngXYDataToCountTrack As Long
    Dim intDataDiscardValue As Integer, strCurrentTask As String
    
    Const MaxLinesCount = 32000
    
    Dim XYDataToPlot() As usrXYData, XYDataToPlotCount As Long
    
    Dim HighlightIndex As Long
    Dim dblPreviousMinimum As Double
    Dim lngChunkSize As Long, lngMinimumValIndex As Long, lngMaximumValIndex As Long
    
    Dim StartXValIndex As Long, EndXValIndex As Long
    Dim dblWork As Double

    Dim intDynamicObjectOffset As Integer, lngDataSetLineColor As Long
    Dim dblMinimumIntensity As Double, dblMaximumIntensity As Double
    
    ' Need to determine the correct scaling values if autoscaling the y-axis
    ' Only works if the data is sorted in the x direction
    If ThesePlotOptions.AutoScaleY Or ThesePlotOptions.FixYAxisMinimumAtZero Then
        ' Find the minimum and maximum y intensities for all data sets within the range of x data being shown
        dblMaximumIntensity = LowestValueForDoubleDataType
        dblMinimumIntensity = HighestValueForDoubleDataType
        For intDataSetIndex = 0 To intDataSetsLoaded - 1
            With ThisXYDataArray(intDataSetIndex)
                If .XYDataListCount > 0 Then
                    ' Step through .XYDataList and find index of the start x value
                    For lngIndex = 1 To .XYDataListCount
                        If .XYDataList(lngIndex).XVal >= ThesePlotOptions.Scaling.x.ValStart.Val Then
                            StartXValIndex = lngIndex
                            Exit For
                        End If
                    Next lngIndex
                
                    ' Step through .XYDataList and find index of the end x value
                    For lngIndex = .XYDataListCount To 1 Step -1
                        If .XYDataList(lngIndex).XVal <= ThesePlotOptions.Scaling.x.ValEnd.Val Then
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
                End If
            End With
        Next intDataSetIndex
        
        If dblMaximumIntensity <= LowestValueForDoubleDataType Then dblMaximumIntensity = 1
        If dblMinimumIntensity >= HighestValueForDoubleDataType Then dblMinimumIntensity = 0
        
        If ThesePlotOptions.FixYAxisMinimumAtZero Then
            ' Fix Y axis range at zero
            ThesePlotOptions.Scaling.Y.ValStart.Val = 0
        Else
            ThesePlotOptions.Scaling.Y.ValStart.Val = dblMinimumIntensity
        End If
        
        If ThesePlotOptions.AutoScaleY Then
            ThesePlotOptions.Scaling.Y.ValEnd.Val = dblMaximumIntensity
        End If
    End If

    strCurrentTask = LookupMessage(1130)
    
    lngSecondsElapsedAtStart = Timer()

    intDynamicObjectOffset = 0
    For intDataSetIndex = 0 To intDataSetsLoaded - 1
        lngDataSetLineColor = GetPlotSeriesColor(intDataSetIndex)
        
        If ThisXYDataArray(intDataSetIndex).XYDataListCount > 0 Then
            
            ' Make sure the Y Data range is defined
            CheckYDataRange ThisXYDataArray(intDataSetIndex).XYDataList, ThisXYDataArray(intDataSetIndex).XYDataListCount, ThesePlotOptions, intDataSetIndex
                
            With ThisXYDataArray(intDataSetIndex)
                
                ' Determine the location in the parent frame of the bottom of the Plot
                PlotBottom = ThesePlotOptions.PlotTop + ThesePlotOptions.PlotHeight
                
                StartXValIndex = 1
                EndXValIndex = .XYDataListCount
                
                ' Record the X scale data limits
                ThesePlotOptions.DataLimits(intDataSetIndex).x.ValStart.Val = .XYDataList(StartXValIndex).XVal
                ThesePlotOptions.DataLimits(intDataSetIndex).x.ValEnd.Val = .XYDataList(EndXValIndex).XVal
                
                ' Initialize .Scaling.x.ValStart.Val and .Scaling.y.ValStart.Val if necessary
                If ThesePlotOptions.Scaling.x.ValStart.Val = 0 And ThesePlotOptions.Scaling.x.ValEnd.Val = 0 Then
                    ThesePlotOptions.Scaling.x.ValStart.Val = .XYDataList(StartXValIndex).XVal
                    ThesePlotOptions.Scaling.x.ValEnd.Val = .XYDataList(EndXValIndex).XVal
                End If
                
                ' Make sure .Scaling.X.ValStart.Val < .Scaling.X.ValEnd
                If ThesePlotOptions.Scaling.x.ValStart.Val > ThesePlotOptions.Scaling.x.ValEnd.Val Then
                    SwapValues ThesePlotOptions.Scaling.x.ValStart.Val, ThesePlotOptions.Scaling.x.ValEnd.Val
                End If
                
                If Not ThesePlotOptions.ZoomOutFull Then
                    ' Step through .XYDataList and find index of the start x value
                    For lngIndex = 1 To .XYDataListCount
                        If .XYDataList(lngIndex).XVal >= ThesePlotOptions.Scaling.x.ValStart.Val Then
                            StartXValIndex = lngIndex
                            Exit For
                        End If
                    Next lngIndex
                
                    ' Step through .XYDataList and find index of the end x value
                    For lngIndex = .XYDataListCount To 1 Step -1
                        If .XYDataList(lngIndex).XVal < ThesePlotOptions.Scaling.x.ValEnd.Val Then
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
                CheckForLongOperation ThisForm, lngSecondsElapsedAtStart, ThesePlotOptions, 2, 10, strCurrentTask
                
                ' Copy the data into XYDataToPlot, reindexing to start at 1 and going to XYDataToPlotCount
                ' Although this uses more memory because it duplicates the data, I soon replace the
                ' raw data values with location positions in twips
                '
                ' In addition, if there is far more data than could be possibly plotted,
                ' I throw away every xth data point, though with a twist
                
                XYDataToPlotCount = EndXValIndex - StartXValIndex + 1
                
                intDataDiscardValue = ThesePlotOptions.ApproximationFactor
                If intDataDiscardValue < 1 Then
                    intDataDiscardValue = 1
                ElseIf intDataDiscardValue > 50 Then
                    intDataDiscardValue = 50
                End If
                
                If XYDataToPlotCount > ThesePlotOptions.PlotWidth / intDataDiscardValue Then
                    ' Throw away some of the data:  Note that CIntSafe will round 2.5 up to 3
                    intKeepEveryXthPoint = CIntSafeDbl(XYDataToPlotCount / (ThesePlotOptions.PlotWidth / intDataDiscardValue))
                    lngChunkSize = intKeepEveryXthPoint * 2
                Else
                    intKeepEveryXthPoint = 1
                End If
                
                ReDim XYDataToPlot(CLngRoundUp(XYDataToPlotCount / intKeepEveryXthPoint) + intDataDiscardValue)
                
                If intKeepEveryXthPoint = 1 Then
                    lngXYDataToCountTrack = 0
                    For lngIndex = StartXValIndex To EndXValIndex
                        lngXYDataToCountTrack = lngXYDataToCountTrack + 1
                        XYDataToPlot(lngXYDataToCountTrack) = .XYDataList(lngIndex)
                        If lngIndex = ThesePlotOptions.IndexToHighlight Then
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
                    XYDataToPlot(lngXYDataToCountTrack) = .XYDataList(StartXValIndex)
                    If StartXValIndex = ThesePlotOptions.IndexToHighlight Then
                        HighlightIndex = lngXYDataToCountTrack
                    End If
                    
                    dblPreviousMinimum = LowestValueForDoubleDataType
                    For lngIndex = StartXValIndex + 1 To EndXValIndex - 1 Step lngChunkSize
                        
                        FindMinimumAndMaximum lngMinimumValIndex, lngMaximumValIndex, .XYDataList(), lngIndex, lngIndex + lngChunkSize
                        
                        ' Check if the maximum value of this pair of points is less than the minimum value of the previous pair
                        ' If it is, the y values of the two points should be exchanged
                        If .XYDataList(lngMaximumValIndex).YVal < dblPreviousMinimum Then
                            ' Update dblPreviousMinimum
                            dblPreviousMinimum = .XYDataList(lngMinimumValIndex).YVal
                            ' Swap minimum and maximum so that maximum gets saved to array first
                            SwapValues lngMinimumValIndex, lngMaximumValIndex
                        Else
                            ' Update dblPreviousMinimum
                            dblPreviousMinimum = .XYDataList(lngMinimumValIndex).YVal
                        End If
                            
                        lngXYDataToCountTrack = lngXYDataToCountTrack + 1
                          XYDataToPlot(lngXYDataToCountTrack) = .XYDataList(lngMinimumValIndex)
                          If lngMinimumValIndex = ThesePlotOptions.IndexToHighlight Then
                            HighlightIndex = lngXYDataToCountTrack
                          End If
                        
                        lngXYDataToCountTrack = lngXYDataToCountTrack + 1
                          XYDataToPlot(lngXYDataToCountTrack) = .XYDataList(lngMaximumValIndex)
                          If lngMaximumValIndex = ThesePlotOptions.IndexToHighlight Then
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
                    XYDataToPlot(lngXYDataToCountTrack) = .XYDataList(EndXValIndex)
                    If EndXValIndex = ThesePlotOptions.IndexToHighlight Then
                        HighlightIndex = lngXYDataToCountTrack
                    End If
                
                End If
                
                ' Set XYDataToPlotCount to the lngXYDataToCountTrack value resulting from the copying
                XYDataToPlotCount = lngXYDataToCountTrack
            
            End With
        
            ' Scale the data vertically according to ThesePlotOptions.height
            ' i.e., replace the actual y values with locations in twips for where the data point belongs on the graph
            ' The new value will range from 0 to ThesePlotOptions.Height
            ' Note that ThesePlotOptions.PlotLeftLargeNumberOffset is computed in ScaleData for the Y axis
            ScaleData ThesePlotOptions, XYDataToPlot(), XYDataToPlotCount, PlotRange(intDataSetIndex).Y, ThesePlotOptions.Scaling.Y, False, ThesePlotOptions.AutoScaleY
            
            ' Now scale the data to twips, ranging 0 to .PlotWidth and 0 to .PlotHeight
            ' X axis
            ScaleData ThesePlotOptions, XYDataToPlot(), XYDataToPlotCount, PlotRange(intDataSetIndex).x, ThesePlotOptions.Scaling.x, True, False
            
            ' Load lines and labels for each XVal
            If ThesePlotOptions.LinesLoadedCount = 0 Then
                ' Only initialize to 1 the first time this sub is called
                ThesePlotOptions.LinesLoadedCount = 1
            End If
            
            ' Limit the total points shown if greater than MaxLinesCount
            If XYDataToPlotCount > MaxLinesCount Then
                ' This code should not be reached since extra data should have been thrown away above
                Debug.Assert False
                XYDataToPlotCount = MaxLinesCount
            End If
            
            ' Load dynamic plot objects as needed
            LoadDynamicPlotObjects ThisForm, ThesePlotOptions, intDynamicObjectOffset + XYDataToPlotCount
            
            If ThesePlotOptions.LinesLoadedCount < intDynamicObjectOffset + XYDataToPlotCount Then
                ' User aborted the process of dynamically loading lines
                ' Need to limit XYDataToPlotCount to the correct value
                XYDataToPlotCount = ThesePlotOptions.LinesLoadedCount - intDynamicObjectOffset
            End If
            
            ' Check to see if Mouse Pointer should be changed to hourglass
            CheckForLongOperation ThisForm, lngSecondsElapsedAtStart, ThesePlotOptions, 3, 10, strCurrentTask
            
            ' Label the axes and add ticks and gridlines
            FormatAxes ThisForm, ThesePlotOptions, PlotBottom, PlotRange(intDataSetIndex)
            
            ' Position the lines and labels
            lngLeftOffset = ThesePlotOptions.PlotLeft + ThesePlotOptions.PlotLeftLargeNumberOffset
            
            strCurrentTask = LookupMessage(1135)
            
            For lngIndex = 1 To XYDataToPlotCount
                With ThisForm
                    If lngIndex Mod 50 = 0 Then
                        ' Check to see if Mouse Pointer should be changed to hourglass
                        CheckForLongOperation ThisForm, lngSecondsElapsedAtStart, ThesePlotOptions, lngIndex, XYDataToPlotCount, strCurrentTask
                        If KeyPressAbortProcess > 1 Then Exit For
                    End If
                    
                    lngLineIndex = lngIndex + intDynamicObjectOffset
                    
                    .linData(lngLineIndex).Visible = True

                    If ThesePlotOptions.PlotTypeCode = cPlotTypeSticks Then
                        ' Plot the data as sticks to zero
                        
                        .linData(lngLineIndex).x1 = lngLeftOffset + XYDataToPlot(lngIndex).XVal
                        .linData(lngLineIndex).x2 = .linData(lngLineIndex).x1
                        .linData(lngLineIndex).y1 = PlotBottom
                        .linData(lngLineIndex).y2 = PlotBottom - XYDataToPlot(lngIndex).YVal
                        If HighlightIndex = lngIndex Then
                            .linData(lngLineIndex).BorderColor = ThesePlotOptions.HighlightColor
                        Else
                            .linData(lngLineIndex).BorderColor = lngDataSetLineColor
                        End If
                    Else
                        ' Plot the data as lines between points
                        If lngIndex < XYDataToPlotCount Then
                            .linData(lngLineIndex).x1 = lngLeftOffset + XYDataToPlot(lngIndex).XVal
                            .linData(lngLineIndex).x2 = lngLeftOffset + XYDataToPlot(lngIndex + 1).XVal
                            .linData(lngLineIndex).y1 = PlotBottom - XYDataToPlot(lngIndex).YVal
                            .linData(lngLineIndex).y2 = PlotBottom - XYDataToPlot(lngIndex + 1).YVal
                            .linData(lngLineIndex).BorderColor = lngDataSetLineColor
                        Else
                            .linData(lngLineIndex).Visible = False
                        End If
                    End If
                End With
                
            Next lngIndex
            
            If KeyPressAbortProcess > 1 Then
                intDynamicObjectOffset = intDynamicObjectOffset + lngIndex
                Exit For
            Else
                intDynamicObjectOffset = intDynamicObjectOffset + XYDataToPlotCount
            End If
            
        End If
    Next intDataSetIndex
    
    With ThisForm
        If intDynamicObjectOffset + 1 < ThesePlotOptions.LinesLoadedCount Then
            If intDynamicObjectOffset < 0 Then lngIndex = 0
            
            ' Hide the other lines and labels
            For lngLineIndex = intDynamicObjectOffset + 1 To ThesePlotOptions.LinesLoadedCount
                ThisForm.linData(lngLineIndex).Visible = False
            Next lngLineIndex
        End If
       
    End With
    
    frmProgress.HideForm
        
End Sub

Private Sub FindMinimumAndMaximum(lngMinimumValIndex As Long, lngMaximumValIndex As Long, ThisXYData() As usrXYData, lngStartIndex As Long, lngStopIndex As Long)
    Dim lngIndex As Long, XYDataPoint As Long
    Dim dblMinimumVal As Double, dblMaximumVal As Double
    
    If lngStopIndex > UBound(ThisXYData()) Then
        lngStopIndex = UBound(ThisXYData())
    End If
    
    lngMinimumValIndex = lngStartIndex
    lngMaximumValIndex = lngStartIndex
    
    dblMaximumVal = ThisXYData(lngStartIndex).YVal
    dblMinimumVal = ThisXYData(lngStartIndex).YVal
    
    For lngIndex = lngStartIndex + 1 To lngStopIndex
        XYDataPoint = ThisXYData(lngIndex).YVal
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

Private Sub FormatAxes(ThisForm As VB.Form, ThesePlotOptions As usrPlotDataOptions, PlotBottom As Long, PlotRange As usrPlotRange)
    Dim intInitialTickIndex As Integer, intInitialTickMinorIndex As Integer
    Dim lngRightOrTopMostPos As Long
    
    With ThisForm
        ' Position the x axis
        .linXAxis.x1 = ThesePlotOptions.PlotLeft + ThesePlotOptions.PlotLeftLargeNumberOffset
        .linXAxis.x2 = ThesePlotOptions.PlotLeft + ThesePlotOptions.PlotWidth + 50
        .linXAxis.y1 = PlotBottom + 50
        .linXAxis.y2 = PlotBottom + 50
        .linXAxis.Visible = ThesePlotOptions.XAxis.Show
        
        ' Position the y axis
        .linYAxis.x1 = ThesePlotOptions.PlotLeft + ThesePlotOptions.PlotLeftLargeNumberOffset
        .linYAxis.x2 = ThesePlotOptions.PlotLeft + ThesePlotOptions.PlotLeftLargeNumberOffset
        .linYAxis.y1 = PlotBottom + 50
        .linYAxis.y2 = ThesePlotOptions.PlotTop - 50
        .linYAxis.Visible = ThesePlotOptions.YAxis.Show
    End With
    
    ' Note: The x and y axes share the same dynamic lines for major and minor tick marks,
    '       gridlines, and labels.  The x axis objects will start with index 1 of each object type
    '       the intInitialTickIndex and intInitialTickMinorIndex values will be modified by sub
    '       FormatThisAxis during the creation of the objects for the x axis so that they will
    '       be the value of the next unused object for operations involving the y axis
    intInitialTickIndex = 0
    intInitialTickMinorIndex = 0
    FormatThisAxis ThisForm, ThesePlotOptions, True, ThesePlotOptions.XAxis, PlotRange.x, intInitialTickIndex, intInitialTickMinorIndex, lngRightOrTopMostPos
    If lngRightOrTopMostPos > ThisForm.linXAxis.x2 Then
        ThisForm.linXAxis.x2 = lngRightOrTopMostPos + 50
    End If
    
    FormatThisAxis ThisForm, ThesePlotOptions, False, ThesePlotOptions.YAxis, PlotRange.Y, intInitialTickIndex, intInitialTickMinorIndex, lngRightOrTopMostPos
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

Private Sub FormatThisAxis(ThisForm As VB.Form, ThesePlotOptions As usrPlotDataOptions, blnXAxis As Boolean, AxisOptions As usrAxisOptions, PlotRangeForAxis As usrPlotRangeAxis, ByRef intInitialTickIndex As Integer, ByRef intInitialTickMinorIndex As Integer, lngRightOrTopMostPos As Long)
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
    
    If blnXAxis Then
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
        
        If blnXAxis Then
            LengthStartPosMinor = LengthStartPos - 150
            LengthEndPosMinor = LengthEndPos
        Else
            LengthStartPosMinor = LengthStartPos + 100
            LengthEndPosMinor = LengthEndPos
        End If
        
        If .ShowGridLinesMajor Then
            If blnXAxis Then
                GridlineEndPos = ThesePlotOptions.PlotTop
            Else
                GridlineEndPos = ThesePlotOptions.PlotLeft + ThesePlotOptions.PlotWidth
            End If
        Else
            GridlineEndPos = LengthEndPos
        End If
    End With
    
    ' Initialize ThesePlotOptions.MajorTicksLoaded and ThesePlotOptions.MinorTicksLoaded if needed
    If ThesePlotOptions.MajorTicksLoaded = 0 Then ThesePlotOptions.MajorTicksLoaded = 1 ' There is always at least 1 loaded
    If ThesePlotOptions.MinorTicksLoaded = 0 Then ThesePlotOptions.MinorTicksLoaded = 1  ' There is always at least 1 loaded

    ' Call FindDigitsInLabelUsingRange to determine the digits in the label and construct the format string
    intDigitsInLabel = FindDigitsInLabelUsingRange(PlotRangeForAxis, intMajorTicksToShow, strFormatString)
    
    ' Each number requires 90 pixels
    intTickLabelWidth = intDigitsInLabel * NUM_TWIPS_PER_DIGIT

    intTickIndexMinorTrack = intInitialTickMinorIndex
    For intTickIndex = intInitialTickIndex + 1 To intInitialTickIndex + intMajorTicksToShow
    
        With AxisOptions
            If ThesePlotOptions.MajorTicksLoaded < intTickIndex Then
                ThesePlotOptions.MajorTicksLoaded = ThesePlotOptions.MajorTicksLoaded + 1
                Load ThisForm.linTickMajor(ThesePlotOptions.MajorTicksLoaded)
                Load ThisForm.lblTick(ThesePlotOptions.MajorTicksLoaded)
                Load ThisForm.linGridline(ThesePlotOptions.MajorTicksLoaded)
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
            .BorderColor = RGB(0, 192, 0)
            .Visible = True
        End With
        
        lngRightOrTopMostPos = LengthEndPos
        If Not blnXAxis Then
            SwapLineCoordinates ThisForm.linTickMajor(intTickIndex)
            SwapLineCoordinates ThisForm.linGridline(intTickIndex)
            Set objTickMajor = ThisForm.linTickMajor(intTickIndex)
        End If
        
        With ThisForm.lblTick(intTickIndex)
            .Width = intTickLabelWidth
            .Caption = Format(ValStart + CDbl((intTickIndex - intInitialTickIndex) - 1) * DeltaVal, strFormatString)
            .Visible = True
            If blnXAxis Then
                .top = objTickMajor.y1 + 50
                .Left = objTickMajor.x1 - .Width / 2
            Else
                .top = objTickMajor.y1 - .Height / 2
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

            ' Only load minor ticks if not on the last major tick
        With AxisOptions
            ' Load minor ticks as needed
            intMinorTicksRequired = intTickIndexMinorTrack + intMinorTicksPerMajorTick
            Do While ThesePlotOptions.MinorTicksLoaded < intMinorTicksRequired
                ThesePlotOptions.MinorTicksLoaded = ThesePlotOptions.MinorTicksLoaded + 1
                Load ThisForm.linTickMinor(ThesePlotOptions.MinorTicksLoaded)
            Loop
        End With
        
        For intTickIndexMinor = 1 To intMinorTicksPerMajorTick
            intTickIndexMinorTrack = intTickIndexMinorTrack + 1
            If blnXAxis Then
                lngAddnlMinorTickPos = objTickMajor.x1 + DeltaPosMinor * intTickIndexMinor
            Else
                lngAddnlMinorTickPos = objTickMajor.y1 + DeltaPosMinor * intTickIndexMinor
            End If
            
            lngRightOrTopMostPos = lngAddnlMinorTickPos
            
            ' Add the minor tick mark
            AddMinorTickmark ThisForm, intTickIndexMinor, intTickIndexMinorTrack, lngAddnlMinorTickPos, LengthStartPosMinor, LengthEndPosMinor, blnXAxis
        Next intTickIndexMinor
    Next intTickIndex
    
    ' See if minor ticks can be added before the first major tick
    If AxisOptions.ShowMinorTicks Then
        If blnXAxis Then
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
            If blnXAxis Then
                If lngAddnlMinorTickPos - Abs(DeltaPosMinor) <= lngAddnlMinorTickStopPos Then Exit Do
                lngAddnlMinorTickPos = lngAddnlMinorTickPos - Abs(DeltaPosMinor)
            Else
                If lngAddnlMinorTickPos + Abs(DeltaPosMinor) >= lngAddnlMinorTickStopPos Then Exit Do
                lngAddnlMinorTickPos = lngAddnlMinorTickPos + Abs(DeltaPosMinor)
            End If
            
            intTickIndexMinor = intTickIndexMinor + 1
            intTickIndexMinorTrack = intTickIndexMinorTrack + 1
            
            ' May need to load more minor ticks
            If ThesePlotOptions.MinorTicksLoaded < intTickIndexMinorTrack Then
                ThesePlotOptions.MinorTicksLoaded = ThesePlotOptions.MinorTicksLoaded + 1
                Load ThisForm.linTickMinor(ThesePlotOptions.MinorTicksLoaded)
            End If

            ' Add the minor tick mark
            AddMinorTickmark ThisForm, intTickIndexMinor, intTickIndexMinorTrack, lngAddnlMinorTickPos, LengthStartPosMinor, LengthEndPosMinor, blnXAxis
        Loop
    End If
    
    With ThisForm
        ' Check for overlapping tick labels
        ' If there is overlap, then show every other or every 5th label
        intKeepEveryXthTick = 1
        If intMajorTicksToShow >= 3 Then
            If blnXAxis Then
                If .lblTick(1).Left + .lblTick(1).Width > .lblTick(3).Left Then intKeepEveryXthTick = 5
            Else
                If .lblTick(intInitialTickIndex + 3).top + .lblTick(intInitialTickIndex + 3).Height > .lblTick(intInitialTickIndex + 1).top Then intKeepEveryXthTick = 5
            End If
        End If
        
        If intKeepEveryXthTick = 1 And intMajorTicksToShow >= 2 Then
            If blnXAxis Then
                If .lblTick(1).Left + .lblTick(1).Width > .lblTick(2).Left Then intKeepEveryXthTick = 2
            Else
                If .lblTick(intInitialTickIndex + 2).top + .lblTick(intInitialTickIndex + 2).Height > .lblTick(intInitialTickIndex + 1).top Then intKeepEveryXthTick = 2
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
        For intTickIndexToHide = intInitialTickIndex + 1 To ThesePlotOptions.MajorTicksLoaded
            .linTickMajor(intTickIndexToHide).Visible = False
            .linGridline(intTickIndexToHide).Visible = False
            .lblTick(intTickIndexToHide).Visible = False
        Next intTickIndexToHide
        
        For intTickIndexToHide = intInitialTickMinorIndex + 1 To ThesePlotOptions.MinorTicksLoaded
            .linTickMinor(intTickIndexToHide).Visible = False
        Next intTickIndexToHide
    End With
    
End Sub

Private Sub AddMinorTickmark(ThisForm As VB.Form, intTickIndexMinor As Integer, intTickIndexMinorTrack As Integer, lngAddnlMinorTickPos As Long, LengthStartPosMinor As Long, LengthEndPosMinor As Long, blnXAxis As Boolean)
    With ThisForm.linTickMinor(intTickIndexMinorTrack)
        .x1 = lngAddnlMinorTickPos
        .x2 = .x1
        .y1 = LengthStartPosMinor + 20
        .y2 = LengthEndPosMinor
        If intTickIndexMinor Mod 5 = 0 Then
            ' Draw the minor tick mark a little longer
            If blnXAxis Then
                .y1 = .y1 + 70
                .y2 = .y2
            Else
                .y1 = .y1 - 50
                .y2 = .y2
            End If
        End If
        .Visible = True
    End With

    If Not blnXAxis Then SwapLineCoordinates ThisForm.linTickMinor(intTickIndexMinorTrack)

End Sub

Private Sub LoadDynamicPlotObjects(ThisForm As VB.Form, ThesePlotOptions As usrPlotDataOptions, XYDataToPlotCount As Long)
    Dim lngLinesLoadedCountPrevious  As Long
    Dim blnShowProgress As Boolean, strCurrentTask As String
    
    strCurrentTask = LookupMessage(1130)
    
    With ThisForm
        ' Load lines for data points if needed
        If ThesePlotOptions.LinesLoadedCount < XYDataToPlotCount Then
            lngLinesLoadedCountPrevious = ThesePlotOptions.LinesLoadedCount
            If Abs(XYDataToPlotCount - lngLinesLoadedCountPrevious) > 100 Then
                blnShowProgress = True
                frmProgress.InitializeForm strCurrentTask, 0, Abs(XYDataToPlotCount - lngLinesLoadedCountPrevious)
                frmProgress.ToggleAlwaysOnTop True
            End If
            
            Do While ThesePlotOptions.LinesLoadedCount < XYDataToPlotCount
                If blnShowProgress Then
                    If ThesePlotOptions.LinesLoadedCount Mod 10 = 0 Then
                        frmProgress.UpdateProgressBar Abs(ThesePlotOptions.LinesLoadedCount - lngLinesLoadedCountPrevious) + 1
                        If KeyPressAbortProcess > 1 Then Exit Do
                    End If
                End If
                
                ThesePlotOptions.LinesLoadedCount = ThesePlotOptions.LinesLoadedCount + 1
                Load .linData(ThesePlotOptions.LinesLoadedCount)
                .linData(ThesePlotOptions.LinesLoadedCount).Visible = False
            Loop
            If blnShowProgress Then frmProgress.HideForm
        End If
    End With
    
End Sub

Public Function GetPlotSeriesColor(intSeriesNumber As Integer) As Long
    ' intSeriesNumber can be 0, 1, 2, or other
    
    Select Case intSeriesNumber
    Case 0: GetPlotSeriesColor = RGB(0, 0, 150)
    Case 1: GetPlotSeriesColor = vbRed
    Case 2: GetPlotSeriesColor = vbGreen
    Case Else: GetPlotSeriesColor = vbMagenta
    End Select

End Function

Private Function RoundToEvenMultiple(ByVal dblValueToRound As Double, ByVal MultipleValue As Double, ByVal blnRoundUp As Boolean) As Double
    Dim intLoopCount As Integer
    Dim strWork As String, dblWork As Double
    Dim lngExponentValue As Long
    
    ' Find the exponent of MultipleValue
    strWork = Format(MultipleValue, "0E+000")
    lngExponentValue = CIntSafe(Right(strWork, 4))
    
    intLoopCount = 0
    Do While Trim(Str(dblValueToRound / MultipleValue)) <> Trim(Str(Round(dblValueToRound / MultipleValue, 0)))
        dblWork = dblValueToRound / 10 ^ (lngExponentValue)
        dblWork = Format(dblWork, "0")
        dblWork = dblWork * 10 ^ (lngExponentValue)
        If blnRoundUp Then
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

Private Sub ScaleData(ThesePlotOptions As usrPlotDataOptions, XYDataToPlot() As usrXYData, XYDataToPlotCount As Long, ByRef ThisAxisRange As usrPlotRangeAxis, ThisAxisScaling As usrPlotRangeAxis, blnXAxis As Boolean, blnAutoScaleAxis As Boolean)

    Dim dblMinimumIntensity As Double
    Dim lngIndex As Long, dblDataPoint As Double
    Dim dblValRange As Double, dblDeltaScaler As Double, ThisAxisWindowLength As Long
    Dim strFormatString As String, intDigitsInLabel As Integer
    
    ' First step through valid data and find the minimum YVal value
    ' Necessary to check if negative
    dblMinimumIntensity = HighestValueForDoubleDataType
    For lngIndex = 1 To XYDataToPlotCount
        If blnXAxis Then
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
    
    If blnAutoScaleAxis Then
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
            If blnXAxis Then
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
    
    ' Scale the data according to ThesePlotOptions.height
    
    If blnXAxis Then
        ThisAxisWindowLength = ThesePlotOptions.PlotWidth
    Else
        ThisAxisWindowLength = ThesePlotOptions.PlotHeight
    End If
        
    dblDeltaScaler = CDbl(ThisAxisWindowLength) / CDbl(dblValRange)
    
    For lngIndex = 1 To XYDataToPlotCount
        If blnXAxis Then
            dblDataPoint = XYDataToPlot(lngIndex).XVal
        Else
            dblDataPoint = XYDataToPlot(lngIndex).YVal
        End If
        
        dblDataPoint = Format((dblDataPoint - ThisAxisScaling.ValStart.Val) * dblDeltaScaler, "0")
        
        If dblDataPoint > ThisAxisWindowLength Then
            dblDataPoint = ThisAxisWindowLength
        Else
            If dblDataPoint < 0 Then
                dblDataPoint = 0
            End If
        End If
        
        If blnXAxis Then
            XYDataToPlot(lngIndex).XVal = dblDataPoint
        Else
            XYDataToPlot(lngIndex).YVal = dblDataPoint
        End If

    Next lngIndex
    
    If Not blnXAxis Then
        ' Need to recompute .PlotLeftLargeNumberOffset
        ' Call FindDigitsInLabelUsingRange to determine the digits in the label and construct the format string
        intDigitsInLabel = FindDigitsInLabelUsingRange(ThisAxisRange, ThesePlotOptions.YAxis.MajorTicksToShow, strFormatString)
        
        ' Store in .PlotLeftLargeNumberOffset
        ThesePlotOptions.PlotLeftLargeNumberOffset = intDigitsInLabel * NUM_TWIPS_PER_DIGIT
    End If
    
    ' Record the current plot range for future reference when zooming
    If blnXAxis Then
        ThisAxisRange.ValStart.Pos = ThesePlotOptions.PlotLeft + ThesePlotOptions.PlotLeftLargeNumberOffset
        ThisAxisRange.ValEnd.Pos = ThesePlotOptions.PlotLeft + ThesePlotOptions.PlotLeftLargeNumberOffset + ThesePlotOptions.PlotWidth
    Else
        ThisAxisRange.ValStart.Pos = ThesePlotOptions.PlotTop + ThesePlotOptions.PlotHeight
        ThisAxisRange.ValEnd.Pos = ThesePlotOptions.PlotTop     ' equivalent to: PlotBottom - CLng((dblValRange - ThisAxisScaling.ValStart) * dblDeltaScaler)
    End If
    
End Sub

Public Sub ShellSortXYData(ByRef XYData() As usrXYData, ByRef XYDataPointerArray() As Long, ByVal lngLowIndex As Long, ByVal lngHighIndex As Long)
    ' Sort the data by XYData().XVal
    ' Rather than sorting the data itself, sorts a pointer array to the data
    
    Dim lngCount As Long
    Dim lngIncrement As Long
    Dim lngIndex As Long
    Dim lngIndexCompare As Long
    Dim lngPointerSwap As Long

    ' sort XYDataPointerArray[lngLowIndex..lngHighIndex]

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
            lngPointerSwap = XYDataPointerArray(lngIndex)
            For lngIndexCompare = lngIndex - lngIncrement To lngLowIndex Step -lngIncrement
                ' Use <= to sort ascending; Use > to sort descending
                If XYData(XYDataPointerArray(lngIndexCompare)).XVal <= XYData(lngPointerSwap).XVal Then Exit For
                XYDataPointerArray(lngIndexCompare + lngIncrement) = XYDataPointerArray(lngIndexCompare)
            Next lngIndexCompare
            XYDataPointerArray(lngIndexCompare + lngIncrement) = lngPointerSwap
        Next lngIndex
        lngIncrement = lngIncrement \ 3
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
Public Function XYPosToValue(ThisPos As Long, ThisRange As usrPlotRangeAxis) As Double
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

Public Function XYValueToPos(ThisValue As Double, ThisRange As usrPlotRangeAxis, IsDeltaValue As Boolean) As Double
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


