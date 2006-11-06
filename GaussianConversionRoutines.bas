Attribute VB_Name = "GaussianConversionRoutines"
Option Explicit

Public Type usrXYData
    XVal As Double          ' Note: Single data type is -3.402823E38 to 3.402823E38
    YVal As Double
End Type

Public Type usrXYDataSet
    XYDataList() As usrXYData                 ' Index 1 to XYDataListCount
    XYDataListCount As Long
    XYDataListCountDimmed As Long
End Type

Private Sub CheckDynamicXYData(ByRef ThisXYDataList() As usrXYData, ThisXYDataListCount As Long, ByRef ThisXYDataListCountDimmed As Long, Optional lngIncrement As Long = 100)
    If ThisXYDataListCount > ThisXYDataListCountDimmed Then
        ThisXYDataListCountDimmed = ThisXYDataListCountDimmed + lngIncrement
        If ThisXYDataListCountDimmed < ThisXYDataListCount Then
            ThisXYDataListCountDimmed = ThisXYDataListCount
        End If
        ReDim Preserve ThisXYDataList(ThisXYDataListCountDimmed)
    End If
End Sub

Private Sub CheckForLongOperation(ThisForm As VB.Form, ByRef blnLongOperationsRequired As Boolean, lngSecondsElapsedAtStart As Long, lngCurrentIteration As Long, lngTotalIterations As Long, strCurrentTask As String)
    ' Checks to see if the current value of Timer() is greater than lngSecondsElapsedAtStart + 1
    '  If it is then blnLongOperationsRequired is turned on and the pointer is changed to an hourglass
    ' Furthermore, if over 2 seconds have elapsed, then a progress box is shown
    Dim lngSecElapsedSinceOperationStart As Long
    
    lngSecElapsedSinceOperationStart = Timer() - lngSecondsElapsedAtStart
    
    If lngSecElapsedSinceOperationStart >= 1 Or blnLongOperationsRequired Then
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

Public Sub ConvertStickDataToGaussian2DArray(ThisForm As VB.Form, ByRef dblXVals() As Double, ByRef dblYVals() As Double, ByRef lngDataCount As Long, ByVal lngResolution As Long, ByVal lngResolutionMass As Long, ByVal intQualityFactor As Integer)
    ' dblXVals() and dblYVals() are parallel arrays, 0-based (thus ranging from 0 to lngDataCount-1)
    ' The arrays should contain stick data
    ' The original data in the arrays will be replaced with Gaussian peaks in place of each "stick"
    ' Note: Assumes dblXVals() is sorted in the x direction
    
    Const MAX_DATA_POINTS = 100000
    Const MASS_PRECISION = 7
    
    Dim lngDataIndex As Long, lngMidPointIndex As Long
    Dim lngStickIndex As Long, DeltaX As Double
    
    Dim dblXValRange As Double, dblXValWindowRange As Double, dblRangeWork As Double
    Dim dblMinimalXValOfWindow As Double, dblMinimalXValSpacing As Double
    Dim blnSearchForMinimumXVal As Boolean
    
    Dim dblXOffSet As Double, dblSigma As Double
    Dim lngExponentValue As Long
    Dim lngSecondsElapsedAtStart As Long, strCurrentTask As String
    
    Dim XYSummation() As usrXYData                                  ' 0-based array
    Dim XYSummationCount As Long, XYSummationCountDimmed As Long
    Dim lngSummationIndex As Long, lngMinimalSummationIndex As Long
    
    Dim DataToAdd() As usrXYData                                    ' 0-based array
    Dim lngDataToAddCount As Long, blnAppendNewData As Boolean
    Dim ThisDataPoint As usrXYData
    
    Static blnLongOperationsRequired As Boolean
    
On Error GoTo ConvertStickDataToGaussianErrorHandler

    If lngDataCount <= 0 Then Exit Sub
    
    lngSecondsElapsedAtStart = Timer()
    
    ' Initialize XYSummation
    XYSummationCount = 0
    XYSummationCountDimmed = 100
    ReDim XYSummation(XYSummationCountDimmed)
    
    ' Determine the data range for dblXVals() and dblYVals()
    dblXValRange = dblXVals(lngDataCount - 1) - dblXVals(0)
    
    If lngResolution < 1 Then lngResolution = 1
    If intQualityFactor < 1 Or intQualityFactor > 75 Then intQualityFactor = 50
        
    ' Compute DeltaX using .lngResolution and .lngResolutionMass
    ' Do not allow the DeltaX to be so small that the total points required > MAX_DATA_POINTS
    DeltaX = lngResolutionMass / lngResolution / intQualityFactor
    ' Make sure DeltaX is a reasonable number
    DeltaX = RoundToMultipleOf10(DeltaX)
    
    If DeltaX = 0 Then DeltaX = 1
    
    ' Set the Window Range to 1/10 the magnitude of the midpoint x value
    dblRangeWork = dblXVals(0) + dblXValRange / 2
    dblRangeWork = RoundToMultipleOf10(dblRangeWork, lngExponentValue)
    
    dblSigma = (lngResolutionMass / lngResolution) / Sqr(5.54)
    
    ' Set the window range (the xvalue window width range) to calculate the Gaussian representation for each data point
    ' The width at the base of a peak is 4 dblSigma
    ' Use a width of 2 * 6 dblSigma
    dblXValWindowRange = 2 * 6 * dblSigma
    
    If dblXValRange / DeltaX > MAX_DATA_POINTS Then
        ' Delta x is too small; change to a reasonable value
        ' This isn't a bug, but it may mean one of the default settings is inappropriate
        Debug.Assert False
        DeltaX = dblXValRange / MAX_DATA_POINTS
    End If
    
    lngDataToAddCount = CLng(dblXValWindowRange / DeltaX)
    
    ' Make sure lngDataToAddCount is odd
    If CSng(lngDataToAddCount) / 2! = Round(CSng(lngDataToAddCount) / 2!, 0) Then
        lngDataToAddCount = lngDataToAddCount + 1
    End If
    
    ReDim DataToAdd(lngDataToAddCount)
    lngMidPointIndex = (lngDataToAddCount + 1) / 2 - 1          ' Note that DataToAdd() is 0-based
    
    ' Compute the Gaussian data for each point in dblXVals()
    strCurrentTask = LookupMessage(1130)
    
    For lngStickIndex = 0 To lngDataCount - 1
        If lngStickIndex Mod 25 = 0 Then
            CheckForLongOperation ThisForm, blnLongOperationsRequired, lngSecondsElapsedAtStart, lngStickIndex, lngDataCount, strCurrentTask
            If KeyPressAbortProcess > 1 Then Exit For
        End If
        
        ' Search through XYSummation to determine the index of the smallest XValue with which
        '   data in DataToAdd could be combined
        lngMinimalSummationIndex = 0
        
        dblMinimalXValOfWindow = dblXVals(lngStickIndex) - (lngMidPointIndex) * DeltaX
        
        blnSearchForMinimumXVal = True
        If XYSummationCount > 0 Then
            If dblMinimalXValOfWindow > XYSummation(XYSummationCount - 1).XVal Then
                lngMinimalSummationIndex = XYSummationCount - 1
                blnSearchForMinimumXVal = False
            End If
        End If
        
        If blnSearchForMinimumXVal Then
            If XYSummationCount <= 0 Then
                lngMinimalSummationIndex = 0
            Else
                For lngSummationIndex = 0 To XYSummationCount - 1
                    If XYSummation(lngSummationIndex).XVal >= dblMinimalXValOfWindow Then
                        lngMinimalSummationIndex = lngSummationIndex - 1
                        If lngMinimalSummationIndex < 0 Then lngMinimalSummationIndex = 0
                        Exit For
                    End If
                Next lngSummationIndex
                If lngSummationIndex >= XYSummationCount Then
                    lngMinimalSummationIndex = XYSummationCount - 1
                End If
            End If
        End If
        
        ' Construct the Gaussian representation for this Data Point
        ThisDataPoint.XVal = dblXVals(lngStickIndex)
        ThisDataPoint.YVal = dblYVals(lngStickIndex)
        
        ' Round ThisDataPoint.XVal to the nearest DeltaX
        ' If .XVal is not an even multiple of DeltaX then bump up .XVal until it is
        ThisDataPoint.XVal = RoundToEvenMultiple(ThisDataPoint.XVal, DeltaX, True)
        
        For lngDataIndex = 0 To lngDataToAddCount - 1
            ' Equation for Gaussian is: Amplitude * Exp[ -(x - mu)^2 / (2*dblSigma^2) ]
            '        Use lngDataIndex, .YVal, and DeltaX
            dblXOffSet = (lngMidPointIndex - lngDataIndex) * DeltaX
            DataToAdd(lngDataIndex).XVal = ThisDataPoint.XVal - dblXOffSet
            DataToAdd(lngDataIndex).YVal = ThisDataPoint.YVal * Exp(-(dblXOffSet) ^ 2 / (2 * dblSigma ^ 2))
        Next lngDataIndex
        
        ' Now merge DataToAdd into XYSummation
        ' XValues in DataToAdd and those in XYSummation have the same DeltaX value
        ' The XValues in DataToAdd might overlap partially with those in XYSummation
        
        lngDataIndex = 0
        ' First, see if the first XValue in DataToAdd is larger than the last XValue in XYSummation
        If XYSummationCount <= 0 Then
            blnAppendNewData = True
        ElseIf DataToAdd(lngDataIndex).XVal > XYSummation(XYSummationCount - 1).XVal Then
            blnAppendNewData = True
        Else
            blnAppendNewData = False
            ' Step through XYSummation() starting at lngMinimalSummationIndex, looking for
            '   the index to start combining data at
            For lngSummationIndex = lngMinimalSummationIndex To XYSummationCount - 1
                If Round(DataToAdd(lngDataIndex).XVal, MASS_PRECISION) <= Round(XYSummation(lngSummationIndex).XVal, MASS_PRECISION) Then
                    
                    ' The following assertion may not be appropriate
                    Debug.Assert Round(XYSummation(lngSummationIndex).XVal, MASS_PRECISION) = Round(DataToAdd(lngDataIndex).XVal, MASS_PRECISION)
                    
                    ' Within Tolerance; start combining the values here
                    Do While lngSummationIndex <= XYSummationCount - 1
                        XYSummation(lngSummationIndex).YVal = XYSummation(lngSummationIndex).YVal + DataToAdd(lngDataIndex).YVal
                        lngSummationIndex = lngSummationIndex + 1
                        lngDataIndex = lngDataIndex + 1
                        If lngDataIndex >= lngDataToAddCount Then
                            ' Successfully combined all of the data
                            Exit Do
                        End If
                    Loop
                    If lngDataIndex < lngDataToAddCount Then
                        ' Data still remains to be added
                        blnAppendNewData = True
                    End If
                    Exit For
                End If
            Next lngSummationIndex
        End If
        
        If blnAppendNewData = True Then
            CheckDynamicXYData XYSummation(), XYSummationCount + lngDataToAddCount - lngDataIndex, XYSummationCountDimmed
            Do While lngDataIndex < lngDataToAddCount
                ThisDataPoint = DataToAdd(lngDataIndex)
                XYSummation(XYSummationCount) = ThisDataPoint
                XYSummationCount = XYSummationCount + 1
                lngDataIndex = lngDataIndex + 1
            Loop
        End If
        
    Next lngStickIndex
    
    
    ' Assure there is a data point at each 1% point along x range (do give better looking plots)
    ' Probably need to add data, but may need to remove some
    dblMinimalXValSpacing = dblXValRange / 100
    
    lngSummationIndex = 0
    Do While lngSummationIndex < XYSummationCount - 1
        If XYSummation(lngSummationIndex).XVal + dblMinimalXValSpacing < XYSummation(lngSummationIndex + 1).XVal Then
            ' Need to insert a data point
            
            XYSummationCount = XYSummationCount + 1
            CheckDynamicXYData XYSummation, XYSummationCount, XYSummationCountDimmed
            For lngDataIndex = XYSummationCount - 1 To lngSummationIndex + 2 Step -1
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
    ThisForm.MousePointer = vbDefault
    
    ' Reset the blnLongOperationsRequired bit
    blnLongOperationsRequired = False
    
    ' Replace the data in dblXVals() and dblYVals() with the XYSummation() data
    ReDim dblXVals(XYSummationCount)
    ReDim dblYVals(XYSummationCount)
    lngDataCount = XYSummationCount
    
    For lngDataIndex = 0 To XYSummationCount - 1
        dblXVals(lngDataIndex) = XYSummation(lngDataIndex).XVal
        dblYVals(lngDataIndex) = XYSummation(lngDataIndex).YVal
    Next lngDataIndex

    Exit Sub
    
ConvertStickDataToGaussianErrorHandler:
    GeneralErrorHandler "ConvertStickDataToGaussian", Err.Number, Err.Description

End Sub

