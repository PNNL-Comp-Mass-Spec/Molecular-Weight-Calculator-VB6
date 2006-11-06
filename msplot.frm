VERSION 5.00
Begin VB.Form frmMsPlot 
   Caption         =   "Plot"
   ClientHeight    =   5460
   ClientLeft      =   165
   ClientTop       =   795
   ClientWidth     =   7665
   HelpContextID   =   3090
   Icon            =   "msplot.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5460
   ScaleWidth      =   7665
   StartUpPosition =   3  'Windows Default
   Tag             =   "13000"
   Begin VB.Frame fraLegend 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   4695
      Begin VB.Line linLegend 
         Index           =   0
         Visible         =   0   'False
         X1              =   360
         X2              =   120
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Label lblLegend 
         Caption         =   "Series 1"
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   4
         Top             =   20
         Width           =   1815
      End
      Begin VB.Line linLegend 
         Index           =   1
         Visible         =   0   'False
         X1              =   2640
         X2              =   2400
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Label lblLegend 
         Caption         =   "Series 2"
         Height          =   195
         Index           =   1
         Left            =   2760
         TabIndex        =   3
         Top             =   20
         Width           =   1815
      End
   End
   Begin VB.Frame fraPlot 
      BorderStyle     =   0  'None
      Height          =   2985
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   4785
      Begin VB.Timer tmrUpdatePlot 
         Interval        =   100
         Left            =   4320
         Top             =   0
      End
      Begin VB.Shape shpZoomBox 
         BorderColor     =   &H000000FF&
         BorderStyle     =   4  'Dash-Dot
         Height          =   855
         Left            =   120
         Top             =   360
         Width           =   855
      End
      Begin VB.Line linData 
         Index           =   1
         Visible         =   0   'False
         X1              =   2160
         X2              =   2160
         Y1              =   600
         Y2              =   1080
      End
      Begin VB.Line linTickMajor 
         DrawMode        =   9  'Not Mask Pen
         Index           =   1
         Visible         =   0   'False
         X1              =   840
         X2              =   840
         Y1              =   240
         Y2              =   720
      End
      Begin VB.Line linXAxis 
         Visible         =   0   'False
         X1              =   1320
         X2              =   600
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line linYAxis 
         Visible         =   0   'False
         X1              =   600
         X2              =   600
         Y1              =   720
         Y2              =   120
      End
      Begin VB.Label lblTick 
         BackStyle       =   0  'Transparent
         Caption         =   "123.45"
         Height          =   200
         Index           =   1
         Left            =   720
         TabIndex        =   1
         Top             =   840
         Visible         =   0   'False
         Width           =   500
      End
      Begin VB.Line linTickMinor 
         DrawMode        =   9  'Not Mask Pen
         Index           =   1
         Visible         =   0   'False
         X1              =   960
         X2              =   960
         Y1              =   480
         Y2              =   720
      End
      Begin VB.Line linGridline 
         BorderColor     =   &H00008000&
         BorderStyle     =   3  'Dot
         DrawMode        =   9  'Not Mask Pen
         Index           =   1
         Visible         =   0   'False
         X1              =   720
         X2              =   720
         Y1              =   240
         Y2              =   720
      End
   End
   Begin VB.Label lblCurrentPos 
      Caption         =   "Loc: 0,0"
      Height          =   225
      Left            =   5280
      TabIndex        =   5
      Top             =   0
      Width           =   2115
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExportData 
         Caption         =   "&Export Data..."
      End
      Begin VB.Menu mnuFileSepBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuPlotType 
         Caption         =   "&Plot Type"
         Begin VB.Menu mnuPlotTypeSticksToZero 
            Caption         =   "&Sticks to Zero"
         End
         Begin VB.Menu mnuPlotTypeGaussianPeaks 
            Caption         =   "&Gaussian Peaks"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuSetResolution 
         Caption         =   "Set Effective &Resolution..."
      End
      Begin VB.Menu mnuOptionsSepBar6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGridLinesXAxis 
         Caption         =   "X Axis Gridlines"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuGridLinesYAxis 
         Caption         =   "Y Axis Gridlines"
         Shortcut        =   ^Y
      End
      Begin VB.Menu mnuTicks 
         Caption         =   "&Ticks to label (approx.)"
         Begin VB.Menu mnuTicksXAxis 
            Caption         =   "&X Axis..."
         End
         Begin VB.Menu mnuTicksYAxis 
            Caption         =   "&Y Axis..."
         End
      End
      Begin VB.Menu mnuPlotQuality 
         Caption         =   "&Plot Quality (affects speed)"
         Begin VB.Menu mnuGaussianQuality 
            Caption         =   "&Gaussian Representation Quality"
         End
         Begin VB.Menu mnuApproximationFactor 
            Caption         =   "&Approximation Factor"
         End
      End
      Begin VB.Menu mnuOptionsSepBar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSetRangeX 
         Caption         =   "Set &X Range..."
      End
      Begin VB.Menu mnuSetRangeY 
         Caption         =   "Set &Y Range..."
      End
      Begin VB.Menu mnuAutoScaleYAxis 
         Caption         =   "&Autoscale Y Axis"
         Checked         =   -1  'True
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuFixMinimumYAtZero 
         Caption         =   "&Fix mimimum Y at zero"
      End
      Begin VB.Menu mnuOptionsSepBar5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuZoomOutToPrevious 
         Caption         =   "&Zoom Out to Previous"
      End
      Begin VB.Menu mnuZoomOutFullScale 
         Caption         =   "Zoom Out to Show All"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuOptionsSepBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCursorMode 
         Caption         =   "&Cursor Mode"
         Begin VB.Menu mnuCursorModeZoom 
            Caption         =   "&Zoom"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuCursorModeMove 
            Caption         =   "&Move"
         End
      End
      Begin VB.Menu mnuShowCurrentPosition 
         Caption         =   "&Show Current Position"
         Checked         =   -1  'True
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuShowLegend 
         Caption         =   "Show &Legend"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuOptionsSepBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuResetToDefaults 
         Caption         =   "Reset to &Default Options"
      End
   End
   Begin VB.Menu mnuZoomOptions 
      Caption         =   "&Zoom Box"
      Begin VB.Menu mnuZoomIn 
         Caption         =   "Zoom &In"
      End
      Begin VB.Menu mnuZoomInHorizontal 
         Caption         =   "Zoom In Horizontal"
      End
      Begin VB.Menu mnuZoomInVertical 
         Caption         =   "Zoom In Vertical"
      End
      Begin VB.Menu mnuZoomSepBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuZoomOut 
         Caption         =   "Zoom &Out"
      End
      Begin VB.Menu mnuZoomOutHorizontal 
         Caption         =   "Zoom Out Horizontal"
      End
      Begin VB.Menu mnuZoomOutVertical 
         Caption         =   "Zoom Out Vertical"
      End
   End
End
Attribute VB_Name = "frmMsPlot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type usrRect
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
End Type

Private Type usrScalingRangeValues
    PlotRangeStretchVal As Double
    StartVal As Double
    EndVal As Double
    LowerLimit As Double
    UpperLimit As Double
End Type

Private ZoomBoxCoords As usrRect
Private PlotOptions As usrPlotDataOptions

Private blnSlidingGraph As Boolean
Private PlotRangeAtMoveStart As usrPlotRange
Private blnResizingWindow As Boolean
Private blnDrawingZoomBox As Boolean, blnZoomBoxDrawn As Boolean

Private intDataSetLoadedCount As Integer                        ' Count of the number of data sets loaded (originally 0); 1 if 1 data set (index 0 of LoadedXYData), 2 if 2 data sets, etc.
Private LoadedXYData(MAX_DATA_SETS) As usrXYDataSet         ' The data to plot; 0-based array, using indices 0 and 1 since MAX_DATA_SETS = 1; However, .XYDataList() is 1=based
Private InitialStickData(MAX_DATA_SETS) As usrXYDataSet     ' 0-based array: If the user submits Stick Data (discrete data points) and requests that the sticks be converted
                                                            '   to a Gaussian representation, then the original, unmodified data is stored here

Private TicksElapsedSinceStart As Long      ' Actually increments 10 times per second rather than 1000 per second since tmrUpdatePlot.Interval = 100
Private TickCountToUpdateAt As Long

Private blnDelayUpdate As Boolean       ' When true, the plot will not update
Private blnUpdatePosition As Boolean
Private CurrentPosX As Double, CurrentPosY As Double

Private Const PLOT_RANGE_HISTORY_COUNT = 20
Private PlotRangeHistory(PLOT_RANGE_HISTORY_COUNT, MAX_DATA_SETS) As usrPlotRange        ' Keeps track of the last 5 plot ranges displayed to allow for undoing

Private Sub EnableDisableZoomMenus(boolEnableMenus As Boolean)

    mnuZoomOptions.Visible = boolEnableMenus

End Sub

Public Sub SetDelayUpdate(blnTurnOnDelay As Boolean, Optional blnZoomOutCompletely As Boolean = True)
    blnDelayUpdate = blnTurnOnDelay
    
    If Not blnTurnOnDelay And frmMsPlot.Visible = True Then
        If blnZoomOutCompletely Then
            ZoomOut blnZoomOutCompletely
        Else
            UpdatePlot True
        End If
    End If
End Sub

Private Sub DetermineCurrentScalingRange(ThisAxisScaling As usrPlotRangeAxis, TheseDataLimits As usrPlotRangeAxis, CurrentScalingRange As usrScalingRangeValues, Optional sngStretchValMultiplier As Single = 0.1)
    
    With CurrentScalingRange
        .PlotRangeStretchVal = (TheseDataLimits.ValEnd.Val - TheseDataLimits.ValStart.Val) * sngStretchValMultiplier
        
        .StartVal = ThisAxisScaling.ValStart.Val - ThisAxisScaling.ValNegativeValueCorrectionOffset
        .EndVal = ThisAxisScaling.ValEnd.Val - ThisAxisScaling.ValNegativeValueCorrectionOffset
        .LowerLimit = TheseDataLimits.ValStart.Val - .PlotRangeStretchVal
        .UpperLimit = TheseDataLimits.ValEnd.Val + .PlotRangeStretchVal
    End With

End Sub

Private Sub EnableDisableExportDataMenu()
    Dim intDataSetIndex As Integer, boolDataPresent As Boolean
    
    boolDataPresent = False
    
    For intDataSetIndex = 0 To intDataSetLoadedCount - 1
        If LoadedXYData(intDataSetIndex).XYDataListCount > 0 Then
            boolDataPresent = True
            Exit For
        End If
    Next intDataSetIndex
    
    mnuExportData.Enabled = boolDataPresent

End Sub

Private Sub EnableDisableMenuCheckmarks()
    Dim boolLinesBetweenPoints As Boolean
    
    With PlotOptions
        boolLinesBetweenPoints = (.PlotTypeCode = cPlotTypeGaussian)
        mnuPlotTypeSticksToZero.Checked = Not boolLinesBetweenPoints
        mnuPlotTypeGaussianPeaks.Checked = boolLinesBetweenPoints
        mnuGridLinesXAxis.Checked = .XAxis.ShowGridLinesMajor
        mnuGridLinesYAxis.Checked = .YAxis.ShowGridLinesMajor
        mnuAutoScaleYAxis.Checked = .AutoScaleY
        
        mnuFixMinimumYAtZero.Checked = .FixYAxisMinimumAtZero
        mnuFixMinimumYAtZero.Enabled = Not mnuAutoScaleYAxis.Checked
        
        mnuSetResolution.Enabled = boolLinesBetweenPoints
        mnuSetRangeY.Enabled = Not (mnuAutoScaleYAxis.Checked)
    End With
    
End Sub

Public Sub ExportData()
    Dim lngIndex As Long, strFilePath As String, strOutput As String
    Dim intDataSetIndex As Integer, lngMaxDataListCount As Long
    Dim OutFileNum As Integer
    
    On Error GoTo WriteProblem
    
    ' 1520 = Data Files, 1525 = .csv
    strFilePath = SelectFile(frmMsPlot.hwnd, "Save File", gLastFileOpenSaveFolder, True, "", ConstructFileDialogFilterMask(LookupMessage(1520), LookupMessage(1525)), 1)
    If Len(strFilePath) = 0 Then
        ' No file selected (or other error)
        Exit Sub
    End If
    
    OutFileNum = FreeFile()
    Open strFilePath For Output As #OutFileNum
    
    If intDataSetLoadedCount > 1 Then
        strOutput = ""
        For intDataSetIndex = 0 To intDataSetLoadedCount - 1
            strOutput = strOutput & "Data Set " & Trim(CStr(intDataSetIndex + 1)) & " X" & "," & "Data Set " & Trim(CStr(intDataSetIndex + 1)) & " Y"
            If intDataSetIndex < intDataSetLoadedCount - 1 Then strOutput = strOutput & ","
        Next intDataSetIndex
        Print #OutFileNum, strOutput
    End If

    ' Determine maximum .XyDataListCount value
    lngMaxDataListCount = 0
    For intDataSetIndex = 0 To intDataSetLoadedCount - 1
        If LoadedXYData(intDataSetIndex).XYDataListCount > lngMaxDataListCount Then
            lngMaxDataListCount = LoadedXYData(intDataSetIndex).XYDataListCount
        End If
    Next intDataSetIndex
    
    For lngIndex = 1 To lngMaxDataListCount
        strOutput = ""
        For intDataSetIndex = 0 To intDataSetLoadedCount - 1
            If lngIndex <= LoadedXYData(intDataSetIndex).XYDataListCount Then
                With LoadedXYData(intDataSetIndex).XYDataList(lngIndex)
                    strOutput = strOutput & .XVal & "," & .YVal
                End With
            Else
                strOutput = strOutput & ","
            End If
            If intDataSetIndex < intDataSetLoadedCount - 1 Then strOutput = strOutput & ","
        Next intDataSetIndex
        Print #OutFileNum, strOutput
    Next lngIndex
    
    Close #OutFileNum

    Exit Sub
    
WriteProblem:
    MsgBox LookupMessage(1030) & ": " & strFilePath
    
End Sub

Private Function FixUpCoordinates(TheseCoords As usrRect) As usrRect
    Dim FixedCoords As usrRect

    FixedCoords = TheseCoords
    
    With FixedCoords
        If .x1 > .x2 Then
            SwapValues .x1, .x2
        End If
        If .y1 < .y2 Then
            SwapValues .y1, .y2
        End If
    End With
    
    FixUpCoordinates = FixedCoords
    
End Function

Private Sub HidePlotForm()
    frmFragmentationModelling.SetUpdateSpectrumMode False
    HideFormShowMain Me
End Sub
Private Sub HideZoomBox(Button As Integer, boolPerformZoom As Boolean)
    
    EnableDisableZoomMenus False
    blnZoomBoxDrawn = False
    blnDrawingZoomBox = False
    
    If shpZoomBox.Visible = False Then
        Exit Sub
    End If
    
    shpZoomBox.Visible = False
    
    If Button = vbLeftButton Then
        If boolPerformZoom Then
            PerformZoom
        End If
    End If
    
End Sub

Private Sub InitializeDummyData(intDataType As Integer)
    ' intDataType can be 0: continuous sine wave
    '                    1: stick data (only 20 points)
    '                    2: stick data (1000's of points, mostly zero, with a few spikes)

    Dim ThisXYDataSet As usrXYDataSet
    Dim x As Long, sngOffset As Single
    
    Const PI = 3.14159265359
    Const DegToRadiansMultiplier = PI / 180 / 10
    
    Randomize Timer
    
    Select Case intDataType
    Case 1
        With ThisXYDataSet
            .XYDataListCount = 14
            ReDim .XYDataList(.XYDataListCount)
            .XYDataList(1).XVal = 154
            .XYDataList(1).YVal = 79
            .XYDataList(2).XVal = 154.51
            .XYDataList(2).YVal = 25
            .XYDataList(3).XVal = 154.95
            .XYDataList(3).YVal = 15
            .XYDataList(4).XVal = 280.2
            .XYDataList(4).YVal = 60
            .XYDataList(5).XVal = 281.15
            .XYDataList(5).YVal = 20
            .XYDataList(6).XVal = 282.201
            .XYDataList(6).YVal = 10
            .XYDataList(7).XVal = 312
            .XYDataList(7).YVal = 23
            .XYDataList(8).XVal = 312.332
            .XYDataList(8).YVal = 5
            .XYDataList(9).XVal = 312.661
            .XYDataList(9).YVal = 2
            .XYDataList(10).XVal = 500
            .XYDataList(10).YVal = 10
            .XYDataList(11).XVal = 589
            .XYDataList(11).YVal = 102
            .XYDataList(12).XVal = 589.247
            .XYDataList(12).YVal = 72.3
            .XYDataList(13).XVal = 589.523
            .XYDataList(13).YVal = 50.7
            .XYDataList(14).XVal = 589.78
            .XYDataList(14).YVal = 30
        End With
        SetXYData ThisXYDataSet, True, True
    Case 2
        With ThisXYDataSet
            .XYDataListCount = 50000
            ReDim .XYDataList(.XYDataListCount)
            For x = 1 To .XYDataListCount
                .XYDataList(x).XVal = 100 + x / 500
                If x Mod 5000 = 0 Then
                    .XYDataList(x).YVal = Rnd(1) * .XYDataListCount / 200 * Rnd(1)
                ElseIf x Mod 3000 = 0 Then
                    .XYDataList(x).YVal = Rnd(1) * .XYDataListCount / 650 * Rnd(1)
                Else
                    .XYDataList(x).YVal = Rnd(1) * 3
                End If
            Next x
        End With
        SetXYData ThisXYDataSet, True, False
    Case Else
        With ThisXYDataSet
            .XYDataListCount = 360! * 100!
            
            ReDim .XYDataList(.XYDataListCount)
            sngOffset = 10
            For x = 1 To .XYDataListCount
                If x Mod 5050 = 0 Then
                    sngOffset = Rnd(1) + 10
                End If
                .XYDataList(x).XVal = CDbl(x) / 1000 - 5
                .XYDataList(x).YVal = sngOffset - Abs((x - .XYDataListCount / 2)) / 10000 + Sin(DegToRadiansMultiplier * x) * Cos(DegToRadiansMultiplier * x / 2) * 1.29967878493163
            Next x
        End With
        SetXYData ThisXYDataSet, False, False
    End Select
    
    PlotOptions.IndexToHighlight = ThisXYDataSet.XYDataListCount / 2
    
    ZoomOut True

End Sub

Private Sub InitializeZoomOrMove(Button As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        With ZoomBoxCoords
            .x1 = x
            .y1 = y
            .x2 = .x1
            .y2 = .y1
        End With
        
        If mnuCursorModeZoom.Checked Then
            ' Zooming
            ' Begin drawing zoom/move box
            shpZoomBox.Visible = True
            blnZoomBoxDrawn = False
            EnableDisableZoomMenus False
            RedrawZoomBox
        Else
            ' Moving (dragging) plot
            ' Will still update the zoom/move box to keep track of how far dragged
            ' but will not draw the box
            shpZoomBox.Visible = False
            blnDrawingZoomBox = False
            blnSlidingGraph = True
            PlotRangeAtMoveStart = PlotRangeHistory(1, 0)
        End If
        
    Else
        If shpZoomBox.Visible = True Then
            ' User clicked outside of zoom box (not left button), hide it
            HideZoomBox Button, False
        Else
            If Button = vbRightButton Then
                ZoomOut False
            End If
        End If
    End If
    
End Sub

Private Sub RespondZoomModeClick(Button As Integer, x As Single, y As Single)
    ' The Zoom Box is visible and the user clicked inside the box
    ' Handle the click differently depending on the location and the button

    If IsZoomBoxLargeEnough(ZoomBoxCoords) Then
        If Button = vbLeftButton Then
            ' Left click inside box: Remove zoom/move box and zoom
            HideZoomBox Button, True
        ElseIf Button = vbRightButton Then
            ' Right click inside box: Show zoom menu
            PopupMenu mnuZoomOptions, vbPopupMenuLeftAlign
        End If
    Else
        ' Clicked inside box but it's too small
        HideZoomBox Button, False
        SetCursorIcon

        If Button = vbLeftButton Then
            ' Left click outside box: Start a new zoom box
            InitializeZoomOrMove Button, x, y
        End If
    End If
    
End Sub

Private Sub ResetMousePointer(Optional Button As Integer = 0)
    If Button = 0 And Me.MousePointer <> vbDefault Then Me.MousePointer = vbDefault
End Sub

Private Sub ResizeZoomBox(Button As Integer, x As Single, y As Single)
    ' For both zooming and moving, resizes the Zoom Box
    
    ' If zooming, then redraw the box
    If shpZoomBox.Visible = True And Button = vbLeftButton Then
        ' Resize zoom/move box
        ZoomBoxCoords.x2 = x
        ZoomBoxCoords.y2 = y
    
        RedrawZoomBox
    End If

    ' If moving, then call the PerformMove sub to initiate a move
    If mnuCursorModeMove.Checked And Button = vbLeftButton Then
        ' Resize zoom/move box
        ZoomBoxCoords.x2 = x
        ZoomBoxCoords.y2 = y
        
        ' Moving plot
        PerformMove
    End If

End Sub

Public Sub SetAutoscaleY(boolEnable As Boolean)
    
    If boolEnable Then
        ' Auto Scale just turned on - replot
        PlotOptions.AutoScaleY = True
        
        ' Call SetFixMinimumYAtZero in order to properly set .FixYAxisMinimumAtZero
        SetFixMinimumYAtZero PlotOptions.FixYAxisMinimumAtZero

        UpdatePlot
    Else
        ' Auto Scale just turned off
        PlotOptions.AutoScaleY = True
        UpdatePlot
        PlotOptions.AutoScaleY = False
    End If
    EnableDisableMenuCheckmarks
    
End Sub

Private Sub SetCursorIcon(Optional x As Single = 0, Optional y As Single = 0)
    If mnuCursorModeMove.Checked Then
        ' Move mode
        ' Me.MouseIcon = imgMouseHand
        ' Me.MousePointer = vbCustom
        Me.MousePointer = vbSizeAll
    Else
        ' Zoom Mode
        If blnZoomBoxDrawn Then
            If IsClickInBox(x, y, ZoomBoxCoords) Then
                If Me.MousePointer <> vbArrowQuestion Then
                    If IsZoomBoxLargeEnough(ZoomBoxCoords) Then
'                        Me.MouseIcon = imgMouseMagnify
'                        Me.MousePointer = vbCustom
                        'Me.MousePointer = vbArrowQuestion
                        Me.MousePointer = vbUpArrow
                    End If
                End If
            Else
                Me.MousePointer = vbCrosshair
            End If
        Else
            Me.MousePointer = vbCrosshair
        End If
    End If
    EnableDisableMenuCheckmarks
End Sub

Public Sub SetCursorMode(boolMoveMode As Boolean)
    
    mnuCursorModeZoom.Checked = Not boolMoveMode
    mnuCursorModeMove.Checked = boolMoveMode

    SetCursorIcon

End Sub

Public Sub SetFixMinimumYAtZero(boolEnable As Boolean)
    With PlotOptions
        ' If sticks to zero plot and autoscaling then fix Y axis minimum at zero
        If .PlotTypeCode = cPlotTypeSticks And .AutoScaleY Then
            .FixYAxisMinimumAtZero = True
        Else
            .FixYAxisMinimumAtZero = boolEnable
        End If
    End With
    
    EnableDisableMenuCheckmarks
    
    UpdatePlot True
    
End Sub

Public Sub SetGridlinesXAxis(boolEnable As Boolean)
    PlotOptions.XAxis.ShowGridLinesMajor = boolEnable
    UpdatePlot False
    EnableDisableMenuCheckmarks
End Sub

Public Sub SetGridlinesYAxis(boolEnable As Boolean)
    PlotOptions.YAxis.ShowGridLinesMajor = boolEnable
    UpdatePlot False
    EnableDisableMenuCheckmarks
End Sub

Public Sub SetLegendVisibility(blnShowLegend As Boolean)
    mnuShowLegend.Checked = blnShowLegend
    fraLegend.Visible = blnShowLegend
    
    fraPlot.Visible = False
    
    PositionControls
    UpdatePlot False
    
    fraPlot.Visible = True
End Sub

Public Sub SetNewRange(blnXAxis As Boolean, boolPromptUserForValues As Boolean, Optional dblNewStartVal As Double = 0, Optional dblNewEndVal As Double = 0)
    Dim CurrentScalingRange As usrScalingRangeValues, ThisDataSetScalingRange As usrScalingRangeValues
    Dim dblDefaultSeparationValue As Double
    Dim intDataSetIndex As Integer
    Dim strFormatString As String
    
    For intDataSetIndex = 0 To intDataSetLoadedCount - 1
        If blnXAxis Then
            DetermineCurrentScalingRange PlotOptions.Scaling.x, PlotOptions.DataLimits(intDataSetIndex).x, ThisDataSetScalingRange, 1000
        Else
            DetermineCurrentScalingRange PlotOptions.Scaling.y, PlotOptions.DataLimits(intDataSetIndex).y, ThisDataSetScalingRange, 1000
        End If
        If intDataSetIndex = 0 Then
           CurrentScalingRange = ThisDataSetScalingRange
        Else
            With CurrentScalingRange
                If ThisDataSetScalingRange.LowerLimit < .LowerLimit Then
                    .LowerLimit = ThisDataSetScalingRange.LowerLimit
                End If
                If ThisDataSetScalingRange.UpperLimit > .UpperLimit Then
                    .UpperLimit = ThisDataSetScalingRange.UpperLimit
                End If
                If ThisDataSetScalingRange.StartVal < .StartVal Then
                    .StartVal = ThisDataSetScalingRange.StartVal
                End If
                If ThisDataSetScalingRange.EndVal > .EndVal Then
                    .EndVal = ThisDataSetScalingRange.EndVal
                End If
            End With
        End If
    Next intDataSetIndex
    
    With CurrentScalingRange
        If .StartVal = 0 And .EndVal = 0 Then
            .StartVal = .LowerLimit
            .EndVal = .UpperLimit
        End If
    End With
    
    If boolPromptUserForValues Or (dblNewStartVal = 0 And dblNewEndVal = 0) Then
            
        With PlotRangeHistory(1, 0)
            If blnXAxis Then
                strFormatString = ConstructFormatString(Abs(.x.ValEnd.Val - .x.ValStart.Val) / 100)
            Else
                strFormatString = ConstructFormatString(Abs(.y.ValEnd.Val - .y.ValStart.Val) / 100)
            End If
        End With
            
        With CurrentScalingRange
            .StartVal = Format(.StartVal, strFormatString)
            .EndVal = Format(.EndVal, strFormatString)
        End With
         
        With frmSetValue
            .Caption = LookupMessage(1040)
            .lblStartVal.Caption = LookupMessage(1050)
            .txtStartVal = CurrentScalingRange.StartVal
            .lblEndVal.Caption = LookupMessage(1055)
            .txtEndVal = CurrentScalingRange.EndVal
            
            ' Round dblDefaultSeparationValue to nearest 1, 2, or 5 (or multiple of 10 thereof)
            dblDefaultSeparationValue = RoundToMultipleOf10((CurrentScalingRange.EndVal - CurrentScalingRange.StartVal) / 10)
            
            .SetLimits True, CurrentScalingRange.LowerLimit, CurrentScalingRange.UpperLimit, dblDefaultSeparationValue
            
            If blnXAxis Then
                .Caption = LookupMessage(1060)
            Else
                .Caption = LookupMessage(1065)
            End If
        
            .Show vbModal
        End With
        
        If UCase(frmSetValue.lblHiddenStatus) <> "OK" Then Exit Sub
        
        ' Set New Range
        With frmSetValue
            If IsNumeric(.txtStartVal) Then dblNewStartVal = CDbl(.txtStartVal)
            If IsNumeric(.txtEndVal) Then dblNewEndVal = CDbl(.txtEndVal)
        End With
    End If
    
    ' Set new scaling value for all loaded data sets
    If blnXAxis Then
        PlotOptions.Scaling.x.ValStart.Val = dblNewStartVal
        PlotOptions.Scaling.x.ValEnd.Val = dblNewEndVal
    Else
        PlotOptions.Scaling.y.ValStart.Val = dblNewStartVal
        PlotOptions.Scaling.y.ValEnd.Val = dblNewEndVal
    End If
    
    UpdatePlot True
    
End Sub

Public Sub SetXYDataVia2DArray(NewXYData() As Double, NewXYDataCount As Long, intDataSetIndexToUse As Integer, boolTreatDataAsDiscretePoints As Boolean, Optional blnZoomOutCompletely As Boolean = True)
    ' Assumes NewXYData() is a 2D array with 2 columns
    ' Further, assumes NewXYData() is a 1-based array in the first dimension but 0-based in the second
    
    Dim ThisXYDataSet As usrXYDataSet, lngIndex As Long
    Dim boolConvertStickDataToGaussianRepresentation As Boolean
    
    If intDataSetIndexToUse < 0 Then
        MsgBox "Invalid data set number.  Must be between 0 and " & Trim(Str(MAX_DATA_SETS - 1)) & "  Assuming value of 0", vbExclamation + vbOKOnly, "Error"
        intDataSetIndexToUse = 0
    ElseIf intDataSetIndexToUse > MAX_DATA_SETS - 1 Then
        MsgBox "Invalid data set number.  Must be between 0 and " & Trim(Str(MAX_DATA_SETS - 1)) & "  Assuming value of " & Trim(Str(MAX_DATA_SETS - 1)), vbExclamation + vbOKOnly, "Error"
        intDataSetIndexToUse = MAX_DATA_SETS - 1
    End If
    
    With ThisXYDataSet
        .XYDataListCount = NewXYDataCount
        .XYDataListCountDimmed = NewXYDataCount + 1
        ReDim .XYDataList(.XYDataListCountDimmed)
        
        For lngIndex = 1 To NewXYDataCount
            .XYDataList(lngIndex).XVal = NewXYData(lngIndex, 0)
            .XYDataList(lngIndex).YVal = NewXYData(lngIndex, 1)
        Next lngIndex
    End With
    
    boolConvertStickDataToGaussianRepresentation = (PlotOptions.PlotTypeCode = cPlotTypeGaussian)
    SetXYData ThisXYDataSet, boolTreatDataAsDiscretePoints, intDataSetIndexToUse, boolConvertStickDataToGaussianRepresentation, blnZoomOutCompletely
    
End Sub

Private Sub SetXYData(NewXYData As usrXYDataSet, boolTreatDataAsDiscretePoints As Boolean, intDataSetIndexToUse As Integer, Optional boolConvertStickDataToGaussianRepresentation As Boolean = False, Optional blnZoomOutCompletely As Boolean = True)
    ' intDataSetIndexToUse can be 0 up to MAX_DATA_SETS, indicating which data set to use
    Dim XYDataPointerArray() As Long, XYDataPointerArrayCount As Long
    Dim intDataSetIndex As Integer, lngIndex As Long
    Dim dblCompareVal As Double
    
    If intDataSetIndexToUse < 0 Then
        MsgBox "Invalid data set number.  Must be between 0 and " & Trim(Str(MAX_DATA_SETS - 1)) & "  Assuming value of 0", vbExclamation + vbOKOnly, "Error"
        intDataSetIndexToUse = 0
    ElseIf intDataSetIndexToUse > MAX_DATA_SETS - 1 Then
        MsgBox "Invalid data set number.  Must be between 0 and " & Trim(Str(MAX_DATA_SETS - 1)) & "  Assuming value of " & Trim(Str(MAX_DATA_SETS - 1)), vbExclamation + vbOKOnly, "Error"
        intDataSetIndexToUse = MAX_DATA_SETS - 1
    End If
    
    ' Reset .DataLimits
    With PlotOptions.DataLimits(intDataSetIndexToUse)
        .x.ValStart.Val = 0
        .y.ValEnd.Val = 0
        .y.ValStart.Val = 0
        .y.ValEnd.Val = 0
    End With

    If intDataSetIndexToUse > intDataSetLoadedCount - 1 Then
        ' The number of data sets loaded must be at least intDataSetIndexToUse + 1
        intDataSetLoadedCount = intDataSetIndexToUse + 1
    End If
    
    ' Sort a Pointer Array to NewXYData, then store in order in InitialStickData or LoadedXYData
    ' First, initialize XYDataPointerArray
    XYDataPointerArrayCount = NewXYData.XYDataListCount
    ReDim XYDataPointerArray(XYDataPointerArrayCount + 1)
    For lngIndex = 1 To XYDataPointerArrayCount
        XYDataPointerArray(lngIndex) = lngIndex
    Next lngIndex
    
    ' Now sort
    Me.MousePointer = vbHourglass
    ShellSortXYData NewXYData.XYDataList(), XYDataPointerArray(), 1, XYDataPointerArrayCount
    Me.MousePointer = vbDefault
    
    If boolTreatDataAsDiscretePoints Then
        ' Copy from NewXYData into InitialStickData()
        ReDim InitialStickData(intDataSetIndexToUse).XYDataList(XYDataPointerArrayCount + 1)
        For lngIndex = 1 To XYDataPointerArrayCount
            InitialStickData(intDataSetIndexToUse).XYDataList(lngIndex) = NewXYData.XYDataList(XYDataPointerArray(lngIndex))
        Next lngIndex
        InitialStickData(intDataSetIndexToUse).XYDataListCount = XYDataPointerArrayCount
        
        ' Load New Data into LoadedXYData; accomplished using SetPlotType provided
        '  InitialStickData(intDataSetIndex).XYDataListCount > 0
        SetPlotType boolConvertStickDataToGaussianRepresentation, False
    Else
        InitialStickData(intDataSetIndexToUse).XYDataListCount = 0
        
        ' Copy from NewXYData into LoadedXYData()
        ReDim LoadedXYData(intDataSetIndexToUse).XYDataList(XYDataPointerArrayCount + 1)
        For lngIndex = 1 To XYDataPointerArrayCount
            LoadedXYData(intDataSetIndexToUse).XYDataList(lngIndex) = NewXYData.XYDataList(XYDataPointerArray(lngIndex))
        Next lngIndex
        LoadedXYData(intDataSetIndexToUse).XYDataListCount = XYDataPointerArrayCount
        
        SetPlotType True, False
    End If
    
    If blnZoomOutCompletely Then
        With PlotOptions
            .Scaling = .DataLimits(0)
            
            ' Check other data sets to see if data limits are outside limits in .Scaling
            For intDataSetIndex = 1 To intDataSetLoadedCount - 1
                dblCompareVal = .DataLimits(intDataSetIndexToUse).x.ValStart.Val
                If dblCompareVal < .Scaling.x.ValStart.Val Then
                    .Scaling.x.ValStart.Val = dblCompareVal
                End If
    
                dblCompareVal = .DataLimits(intDataSetIndexToUse).x.ValEnd.Val
                If dblCompareVal < .Scaling.x.ValEnd.Val Then
                    .Scaling.x.ValEnd.Val = dblCompareVal
                End If
    
                dblCompareVal = .DataLimits(intDataSetIndexToUse).y.ValStart.Val
                If dblCompareVal > .Scaling.y.ValStart.Val Then
                    .Scaling.y.ValStart.Val = dblCompareVal
                End If
    
                dblCompareVal = .DataLimits(intDataSetIndexToUse).y.ValEnd.Val
                If dblCompareVal > .Scaling.y.ValEnd.Val Then
                    .Scaling.y.ValEnd.Val = dblCompareVal
                End If
            Next intDataSetIndex
            
        End With
    
    End If
    
    ' Reset the blnLongOperationsRequired bit
    PlotOptions.blnLongOperationsRequired = False
    
    ' Erase all data in PlotRangeHistory()
    ' Since this array is dimensioned using the const PLOT_RANGE_HISTORY_COUNT and MAX_DATA_SETS, it
    '  does not need to be re-dimensioned after erasing
    Erase PlotRangeHistory()
    
    RefreshPlot blnZoomOutCompletely

'''    LoadDataIntoDynaPlot LoadedXYData(intDataSetIndexToUse), intDataSetIndexToUse
    
End Sub

'''Private Sub LoadDataIntoDynaPlot(ThisXYData As usrXYDataSet, intDataSetIndexToUse As Integer)
'''
'''    Dim XData() As Double, YData() As Double
'''    Dim lngIndex As Long
'''
'''    ' No longer using DynaPlot, exit sub
'''    Exit Sub
'''
'''    If ThisXYData.XYDataListCount = 0 Then Exit Sub
'''
'''    With ThisXYData
'''
'''        ReDim XData(.XYDataListCount)
'''        ReDim YData(.XYDataListCount)
'''
'''        For lngIndex = 1 To .XYDataListCount
'''            XData(lngIndex) = .XYDataList(lngIndex).XVal
'''            YData(lngIndex) = .XYDataList(lngIndex).YVal
'''        Next lngIndex
'''
'''    End With
'''
''''''    If dynMSPlot.DataCurves.count < intDataSetIndexToUse + 1 Then
''''''        ' Load data into dynPlot
''''''        dynMSPlot.DataCurves.AddVB "Set " & intDataSetIndexToUse + 1, XData(), YData(), False
''''''    Else
''''''        dynMSPlot.DataCurves.Item(intDataSetIndexToUse).Curve.UpdateDataVB XData(), YData()
''''''    End If
'''
'''End Sub

Public Sub SetGaussianQuality(Optional intQualityFactor As Integer = -1)
    Dim strResponse As String, intNewQualityFactor As Integer
    Dim intDataSetIndex As Integer
    
    If intQualityFactor < 0 Then
        strResponse = InputBox(LookupMessage(1070) & ":", LookupMessage(1072), PlotOptions.GaussianConversion.QualityFactor)
    Else
        strResponse = Trim(CStr(intQualityFactor))
    End If
    
    If IsNumeric(strResponse) Then
        intNewQualityFactor = CIntSafe(strResponse)
        
        If intNewQualityFactor < 1 Or intNewQualityFactor > 50 Then
            intNewQualityFactor = 20
        End If
        
        PlotOptions.GaussianConversion.QualityFactor = intNewQualityFactor
    End If
    
    ' The gaussian representation will be updated to reflect the new settings in SetPlotType
    SetPlotType CBool(PlotOptions.PlotTypeCode = cPlotTypeGaussian)
    
End Sub

Public Sub SetPlotApproximationFactor(Optional intApproximationFactor As Integer = -1)
    Dim strResponse As String, intNewFactor As Integer
    
    If intApproximationFactor < 0 Then
        strResponse = InputBox(LookupMessage(1075) & ":", LookupMessage(1077), PlotOptions.ApproximationFactor)
    Else
        strResponse = Trim(CStr(intApproximationFactor))
    End If
    
    If IsNumeric(strResponse) Then
        intNewFactor = CIntSafe(strResponse)
        
        If intNewFactor < 1 Or intNewFactor > 50 Then
            intNewFactor = 10
        End If
        
        PlotOptions.ApproximationFactor = intNewFactor
    End If
    
    UpdatePlot False

End Sub

Private Sub InitializePlotOptions()
    
    SetLegendVisibility False
    
    With PlotOptions
        .PlotTypeCode = cPlotTypeSticks
        .AutoScaleY = True
        .ApproximationFactor = 10
        
        .XAxis.Show = True
        .XAxis.MajorTicksToShow = 5
        .XAxis.MinorTickMinimumPixelSep = 100
        .XAxis.ShowMinorTicks = True
        .XAxis.ShowGridLinesMajor = False
        
        .YAxis.Show = True
        .YAxis.MajorTicksToShow = 5
        .YAxis.MinorTickMinimumPixelSep = 100
        .YAxis.ShowMinorTicks = True
        .YAxis.ShowGridLinesMajor = False
            
        .GaussianConversion.ResolvingPower = 5000
        .GaussianConversion.XValueOfSpecification = 500
        .GaussianConversion.QualityFactor = 20
        
        .HighlightColor = vbRed
    End With
    
    Dim intIndex As Integer
    For intIndex = 0 To 1
        linLegend(intIndex).BorderColor = GetPlotSeriesColor(intIndex)
        linLegend(intIndex).Visible = True
    Next intIndex
    
    lblLegend(0) = LookupLanguageCaption(13030, "Predicted Ions")
    lblLegend(1) = LookupLanguageCaption(13035, "Loaded Ions")
    
    SetCursorMode False
    SetShowCursorPosition True
    SetLegendVisibility False

End Sub

Public Sub SetPlotType(boolLinesBetweenPoints As Boolean, Optional boolUpdatePlot As Boolean = True)
    Dim intDataSetIndex As Integer
    
    If boolLinesBetweenPoints Then
        PlotOptions.PlotTypeCode = cPlotTypeGaussian
        SetFixMinimumYAtZero False
    Else
        PlotOptions.PlotTypeCode = cPlotTypeSticks
        SetFixMinimumYAtZero True
    End If
    
    mnuPlotQuality.Enabled = boolLinesBetweenPoints
    
    For intDataSetIndex = 0 To intDataSetLoadedCount - 1
        
        ' Reset .DataLimits
        With PlotOptions.DataLimits(intDataSetIndex)
            .x.ValStart.Val = 0
            .y.ValEnd.Val = 0
            .y.ValStart.Val = 0
            .y.ValEnd.Val = 0
        End With
            
        If InitialStickData(intDataSetIndex).XYDataListCount > 0 Then
            ' Stick data is present; need to take action
            If boolLinesBetweenPoints Then
                LoadedXYData(intDataSetIndex) = ConvertStickDataToGaussian(Me, InitialStickData(intDataSetIndex), PlotOptions, intDataSetIndex)
            Else
                LoadedXYData(intDataSetIndex) = InitialStickData(intDataSetIndex)
            End If
        Else
            ' Clear data in LoadedXYData()
            LoadedXYData(intDataSetIndex).XYDataListCount = 0
        End If
    Next intDataSetIndex
    
    If boolUpdatePlot Then UpdatePlot False
    
    EnableDisableMenuCheckmarks
    
End Sub

Public Sub SetResolution(Optional lngNewResolvingPower As Long = -1, Optional dblNewXValResLocation As Double = 500)
    Dim strResponse As String
    
    If lngNewResolvingPower < 1 Then
        With frmSetValue
            .Caption = LookupMessage(1080)
            .lblStartVal.Caption = LookupMessage(1090)
            .txtStartVal = PlotOptions.GaussianConversion.ResolvingPower
            .lblEndVal.Caption = LookupMessage(1100)
            .txtEndVal = PlotOptions.GaussianConversion.XValueOfSpecification
            
            .SetLimits False
            
            .Show vbModal
        End With
        
        If UCase(frmSetValue.lblHiddenStatus) <> "OK" Then Exit Sub
        
        ' Set New Range
        With frmSetValue
            If IsNumeric(.txtStartVal) Then
                lngNewResolvingPower = CLng(.txtStartVal)
            Else
                lngNewResolvingPower = 5000
            End If
            If IsNumeric(.txtEndVal) Then
                dblNewXValResLocation = CDbl(.txtEndVal)
            Else
                dblNewXValResLocation = 500
            End If
        End With
        
    End If
    
    If lngNewResolvingPower < 1 Or lngNewResolvingPower > 1E+38 Then
        lngNewResolvingPower = 5000
    End If
    
    If dblNewXValResLocation < 1 Or dblNewXValResLocation > 1E+38 Then
        dblNewXValResLocation = 500
    End If
    
    PlotOptions.GaussianConversion.ResolvingPower = lngNewResolvingPower
    PlotOptions.GaussianConversion.XValueOfSpecification = dblNewXValResLocation
    
    ' Must recreate Gaussian representation (if Gaussian mode is on)
    SetPlotType CBool(PlotOptions.PlotTypeCode), True

End Sub

Public Sub SetShowCursorPosition(boolEnable As Boolean)
    mnuShowCurrentPosition.Checked = boolEnable
    lblCurrentPos.Visible = boolEnable
    fraPlot.Visible = False
    
    PositionControls
    UpdatePlot False
    
    fraPlot.Visible = True
End Sub

Private Sub AxisTickCountUpdateByUser(blnXAxis As Boolean, intCurrentTickCount As Integer)
    Dim strAxisLetter As String
    Dim strResponse As String, intNewTickCount As Integer
    
    If blnXAxis Then
        strAxisLetter = "X"
    Else
        strAxisLetter = "Y"
    End If
    
    If intCurrentTickCount < 2 Or intCurrentTickCount > 30 Then intCurrentTickCount = 5
    
    strResponse = InputBox(LookupMessage(1110) & " (" & strAxisLetter & ", 2 - 30):", LookupMessage(1115), intCurrentTickCount)
    
    If IsNumeric(strResponse) Then
        intNewTickCount = CIntSafe(strResponse)
        AxisTickCountUpdate blnXAxis, intNewTickCount
    End If
    
    UpdatePlot False

End Sub

Public Sub AxisTickCountUpdate(blnXAxis As Boolean, intNewTickCount As Integer)
    If intNewTickCount < 2 Or intNewTickCount > 30 Then intNewTickCount = 5

    If blnXAxis Then
        PlotOptions.XAxis.MajorTicksToShow = intNewTickCount
    Else
        PlotOptions.YAxis.MajorTicksToShow = intNewTickCount
    End If
End Sub

Private Sub UpdateCurrentPos()
    Dim XValue As Double, YValue As Double
    Dim strNewString As String
    Dim strFormatStringX As String, strFormatStringY As String
    
    If mnuShowCurrentPosition.Checked Then
        With PlotRangeHistory(1, 0)
            XValue = XYPosToValue(CLng(CurrentPosX), .x)
            YValue = XYPosToValue(CLng(CurrentPosY), .y)
            strNewString = ConstructFormatString(Abs(.x.ValEnd.Val - .x.ValStart.Val) / 100)
            strFormatStringX = strNewString
            
            strNewString = ConstructFormatString(Abs(.y.ValEnd.Val - .y.ValStart.Val) / 100)
            strFormatStringY = strNewString
        End With
        
        lblCurrentPos = LookupLanguageCaption(13010, "Loc") & ": " & Format(XValue, strFormatStringX) & ", " & Format(YValue, strFormatStringY)
    End If
    
    blnUpdatePosition = False
End Sub

Private Sub UpdatePlot(Optional boolUpdateHistory As Boolean = True)
    Dim MostRecentPlotRange(MAX_DATA_SETS) As usrPlotRange, intHistoryIndex As Integer
    Dim intDataSetIndex As Integer
    
    If blnDelayUpdate Then Exit Sub
    
    With PlotOptions
        .PlotTop = 100
        .PlotLeft = 300
        .PlotWidth = fraPlot.Width - .PlotLeft - 0
        .PlotHeight = fraPlot.Height - .PlotTop - 700
        If .PlotHeight < 500 Then .PlotHeight = 500
    End With
    
    For intDataSetIndex = 0 To intDataSetLoadedCount - 1
        MostRecentPlotRange(intDataSetIndex) = PlotRangeHistory(1, intDataSetIndex)
    Next intDataSetIndex
    
    ' Hide fraPlot so that it updates faster
    fraPlot.Visible = False
    
    ' Perform the actual update
    DrawPlot Me, PlotOptions, LoadedXYData(), MostRecentPlotRange(), intDataSetLoadedCount
        
    ' Show fraPlot
    fraPlot.Visible = True
        
    If boolUpdateHistory Then
        ' Update the plot range history
        For intHistoryIndex = PLOT_RANGE_HISTORY_COUNT To 2 Step -1
            For intDataSetIndex = 0 To intDataSetLoadedCount - 1
                PlotRangeHistory(intHistoryIndex, intDataSetIndex) = PlotRangeHistory(intHistoryIndex - 1, intDataSetIndex)
            Next intDataSetIndex
        Next intHistoryIndex
    End If
    
    If Me.MousePointer = vbHourglass Then Me.MousePointer = vbDefault
    
    For intDataSetIndex = 0 To intDataSetLoadedCount - 1
        PlotRangeHistory(1, intDataSetIndex) = MostRecentPlotRange(intDataSetIndex)
    Next intDataSetIndex
End Sub

Private Sub UpdatePlotWhenIdle()
    Dim boolUpdateHistory As Boolean
    
    If TickCountToUpdateAt < TicksElapsedSinceStart Then
        ' Update the plot, but don't update the history if sliding
        If blnSlidingGraph Or blnResizingWindow Then
            boolUpdateHistory = False
        Else
            boolUpdateHistory = True
        End If
        UpdatePlot boolUpdateHistory
        
        TickCountToUpdateAt = 0
        blnResizingWindow = False
    End If
    
End Sub

Private Sub ZoomInHorizontal()
    ' Zoom in along the horizontal axis but
    ' Do not change the vertical range
    
    FixUpCoordinates ZoomBoxCoords
    With ZoomBoxCoords
        .y1 = PlotRangeHistory(1, 0).y.ValEnd.Pos
        .y2 = PlotRangeHistory(1, 0).y.ValStart.Pos
    End With
     
    HideZoomBox vbLeftButton, True

End Sub

Private Sub ZoomInVertical()
    ' Zoom in along the horizontal axis but
    ' Do not change the vertical range
    
    FixUpCoordinates ZoomBoxCoords
    With ZoomBoxCoords
        .x1 = PlotRangeHistory(1, 0).x.ValStart.Pos
        .x2 = PlotRangeHistory(1, 0).x.ValEnd.Pos
    End With
     
    HideZoomBox vbLeftButton, True

End Sub

Private Sub ZoomOut(ByVal blnZoomOutCompletely As Boolean)
    Dim intHistoryIndex As Integer, lngIndex As Long, intDataSetIndex As Integer
    Dim dblCompareXVal As Double, dblMinXVal As Double, dblMaxXVal As Double
    Dim dblPlotRangeStretchVal As Double
    Dim boolAutoScaleSaved As Boolean

    If Not blnZoomOutCompletely Then
        ' See if any previous PlotRange data exists in the history
        ' If not then set blnZoomOutCompletely to True
        With PlotRangeHistory(2, 0)
            If .x.ValStart.Pos = 0 And .x.ValEnd.Pos = 0 And .x.ValStart.Pos = 0 And .x.ValEnd.Pos = 0 Then
                ' Most recent saved zoom range is all zeroes -- not usable
               blnZoomOutCompletely = True
            End If
        End With
    End If
    
    If blnZoomOutCompletely Then
        
        With PlotOptions
            boolAutoScaleSaved = .AutoScaleY
            
            ' Override the AutoScaleY option and turn on ZoomOutFull
            .ZoomOutFull = True
            .AutoScaleY = True
            
            ' Initialize .Scaling.x.ValStart.Val and .Scaling.y.ValStart.Val
            ' Examine all loaded data sets to determine minimum and maximum XVal
            For intDataSetIndex = 0 To intDataSetLoadedCount - 1
                If LoadedXYData(intDataSetIndex).XYDataListCount > 0 Then
                    ' Data in LoadedXYData is not necessarily sorted by mass
                    ' Must step through data to determine minimum and maximum XVal
                    If dblMinXVal = 0 Then
                        dblMinXVal = LoadedXYData(intDataSetIndex).XYDataList(1).XVal
                    End If
                    If dblMaxXVal = 0 Then
                        dblMaxXVal = LoadedXYData(intDataSetIndex).XYDataList(LoadedXYData(intDataSetIndex).XYDataListCount).XVal
                    End If
                    
                    For lngIndex = 1 To LoadedXYData(intDataSetIndex).XYDataListCount
                        dblCompareXVal = LoadedXYData(intDataSetIndex).XYDataList(lngIndex).XVal
                        If dblCompareXVal < dblMinXVal Then dblMinXVal = dblCompareXVal
                        If dblCompareXVal > dblMaxXVal Then dblMaxXVal = dblCompareXVal
                    Next lngIndex
                    
                End If
            Next intDataSetIndex
            
            .Scaling.x.ValStart.Val = dblMinXVal
            .Scaling.x.ValEnd.Val = dblMaxXVal
            
            ' Stretch the limits of the plot by 5% of the total range
            dblPlotRangeStretchVal = (.Scaling.x.ValEnd.Val - .Scaling.x.ValStart.Val) * 0.05
            .Scaling.x.ValEnd.Val = .Scaling.x.ValEnd.Val + dblPlotRangeStretchVal
            .Scaling.x.ValStart.Val = .Scaling.x.ValStart.Val - dblPlotRangeStretchVal

        End With
        
        ' Update the plot
        UpdatePlot
        
        ' Reset .AutoScaleY and turn off .ZoomOutFull
        PlotOptions.AutoScaleY = boolAutoScaleSaved
        PlotOptions.ZoomOutFull = False
    Else
        ' Zoom to previous range
        PlotOptions.Scaling = PlotRangeHistory(2, 0)
        
        UpdatePlot False
        
        ' Update the plot range history
        For intHistoryIndex = 2 To PLOT_RANGE_HISTORY_COUNT - 1
            For intDataSetIndex = 0 To intDataSetLoadedCount - 1
                PlotRangeHistory(intHistoryIndex, intDataSetIndex) = PlotRangeHistory(intHistoryIndex + 1, intDataSetIndex)
            Next intDataSetIndex
        Next intHistoryIndex
        
    End If

End Sub

Private Sub ZoomShrink(boolFixHorizontal As Boolean, boolFixVertical As Boolean)
    ' Zoom out, but not completely
    
    Dim lngViewRangePosX As Long, lngViewRangePosY As Long
    Dim lngBoxSizeX As Double, lngBoxSizeY As Double
    Dim lngPosCorrectionFactorX As Long, lngPosCorrectionFactorY As Long
    Dim TheseCoords As usrRect
    
    With PlotRangeHistory(1, 0)
        lngViewRangePosX = .x.ValEnd.Pos - .x.ValStart.Pos
        lngViewRangePosY = .y.ValEnd.Pos - .y.ValStart.Pos
    End With
    
    If lngViewRangePosX = 0 Or lngViewRangePosY = 0 Then Exit Sub
    
    TheseCoords = ZoomBoxCoords
    FixUpCoordinates TheseCoords
    With TheseCoords
        lngBoxSizeX = Abs(.x2 - .x1)
        lngBoxSizeY = Abs(.y2 - .y1)
        If boolFixVertical Then
            .x1 = PlotRangeHistory(1, 0).x.ValStart.Pos
            .x2 = PlotRangeHistory(1, 0).x.ValEnd.Pos
        Else
            If lngBoxSizeX > 0 Then
                lngPosCorrectionFactorX = (CLng(lngViewRangePosX * CDbl(lngViewRangePosX) / CDbl(lngBoxSizeX))) / 2
                .x1 = .x1 - lngPosCorrectionFactorX
                .x2 = .x2 + lngPosCorrectionFactorX
            End If
        End If
        If boolFixHorizontal Then
            .y1 = PlotRangeHistory(1, 0).y.ValEnd.Pos
            .y2 = PlotRangeHistory(1, 0).y.ValStart.Pos
        Else
            If lngBoxSizeY > 0 Then
                lngPosCorrectionFactorY = (CLng(lngViewRangePosY * CDbl(lngViewRangePosY) / CDbl(lngBoxSizeY))) / 2
                .y1 = .y1 - lngPosCorrectionFactorY
                .y2 = .y2 + lngPosCorrectionFactorY
            End If
        End If
    End With
     
    If lngBoxSizeX > 0 And lngBoxSizeY > 0 Then
        ZoomBoxCoords = TheseCoords
        HideZoomBox vbLeftButton, True
    Else
        HideZoomBox vbLeftButton, False
    End If

End Sub

Private Function IsClickInBox(x As Single, y As Single, TheseZoomBoxCoords As usrRect) As Boolean
    Dim FixedCoords As usrRect

    ' Determine if click was inside or outside of zoom box
    FixedCoords = FixUpCoordinates(TheseZoomBoxCoords)
            
    With FixedCoords
        If x >= .x1 And x <= .x2 And _
           y >= .y2 And y <= .y1 Then
            IsClickInBox = True
        Else
            IsClickInBox = False
        End If
    End With

End Function

Private Function IsZoomBoxLargeEnough(TheseCoords As usrRect) As Boolean
    
    With TheseCoords
        ' Don't zoom if box size is less than 150 by 150 twips
        If Abs(.x2 - .x1) >= 150 And Abs(.y2 - .y1) >= 150 Then
            IsZoomBoxLargeEnough = True
        Else
            IsZoomBoxLargeEnough = False
        End If
    End With

End Function

Private Sub PerformMove()

    Dim TheseCoords As usrRect
    Dim intDataSetIndex As Integer
    Dim dblCompareVal As Double, dblMinDefinedXVal As Double, dblMaxDefinedXVal As Double
    Dim DeltaXVal As Double, DeltaYVal As Double
    Dim dblMinAllowableXVal As Double, dblMaxAllowableXVal As Double, MaximumRange As Double
    
    TheseCoords = ZoomBoxCoords
            
    With PlotRangeAtMoveStart
        DeltaXVal = XYPosToValue(TheseCoords.x2, .x) - XYPosToValue(TheseCoords.x1, .x)
        DeltaYVal = XYPosToValue(TheseCoords.y2, .y) - XYPosToValue(TheseCoords.y1, .y)
    End With
    
    PlotOptions.ZoomOutFull = False
    With PlotOptions
        ' First determine minimum and maximum defined x values for all loaded data sets
        dblMinDefinedXVal = .DataLimits(0).x.ValStart.Val
        dblMaxDefinedXVal = .DataLimits(0).x.ValEnd.Val
        For intDataSetIndex = 1 To intDataSetLoadedCount - 1
            dblCompareVal = .DataLimits(intDataSetIndex).x.ValStart.Val
            If dblCompareVal < dblMinDefinedXVal Then dblMinDefinedXVal = dblCompareVal
            
            dblCompareVal = .DataLimits(intDataSetIndex).x.ValEnd.Val
            If dblCompareVal > dblMaxDefinedXVal Then dblMaxDefinedXVal = dblCompareVal
        Next intDataSetIndex
        
        MaximumRange = dblMaxDefinedXVal - dblMinDefinedXVal
        .Scaling.x.ValStart.Val = PlotRangeAtMoveStart.x.ValStart.Val - DeltaXVal
        dblMinAllowableXVal = dblMinDefinedXVal - MaximumRange / 10
        If .Scaling.x.ValStart.Val < dblMinAllowableXVal Then
            .Scaling.x.ValStart.Val = dblMinAllowableXVal
            .Scaling.x.ValEnd.Val = .Scaling.x.ValStart.Val + (PlotRangeAtMoveStart.x.ValEnd.Val - PlotRangeAtMoveStart.x.ValStart.Val)
        Else
            .Scaling.x.ValEnd.Val = PlotRangeAtMoveStart.x.ValEnd.Val - DeltaXVal
        End If
        
        dblMaxAllowableXVal = dblMaxDefinedXVal + MaximumRange / 10
        If .Scaling.x.ValEnd.Val > dblMaxAllowableXVal Then
            .Scaling.x.ValEnd.Val = dblMaxAllowableXVal
            .Scaling.x.ValStart.Val = .Scaling.x.ValEnd.Val - (PlotRangeAtMoveStart.x.ValEnd.Val - PlotRangeAtMoveStart.x.ValStart.Val)
        End If
        .Scaling.y.ValStart.Val = PlotRangeAtMoveStart.y.ValStart.Val - DeltaYVal
        .Scaling.y.ValEnd.Val = PlotRangeAtMoveStart.y.ValEnd.Val - DeltaYVal
        
    End With

    ' By setting TickCountToUpdateAt to a nonzero value (>= TicksElapsedSinceStart), the
    ' move will be performed when TicksElapsedSinceStart reaches TickCountToUpdateAt
    TickCountToUpdateAt = TicksElapsedSinceStart
End Sub

Private Sub PerformZoom()
    Dim TheseCoords As usrRect
    Dim PlotRangeSaved As usrPlotRange
    
    ' Use the numbers stored in PlotRangeSaved to update the PlotOptions with the desired zoom range
    PlotRangeSaved = PlotRangeHistory(1, 0)  ' The most recent plot range
    
    TheseCoords = FixUpCoordinates(ZoomBoxCoords)
            
    If IsZoomBoxLargeEnough(TheseCoords) Then
        PlotOptions.ZoomOutFull = False
        With PlotOptions.Scaling
            .x.ValStart.Val = XYPosToValue(TheseCoords.x1, PlotRangeSaved.x)
            .x.ValEnd.Val = XYPosToValue(TheseCoords.x2, PlotRangeSaved.x)
            .y.ValStart.Val = XYPosToValue(TheseCoords.y1, PlotRangeSaved.y)
            .y.ValEnd.Val = XYPosToValue(TheseCoords.y2, PlotRangeSaved.y)
        End With
        
        UpdatePlot
    
    End If

End Sub

Private Sub PositionControls()
    Dim PlotHeight As Long
    Dim lngSuggestedPos As Long, lngMinimumPos As Long
    
    With fraLegend
        .Top = 0
        .Left = 720
        .Width = 4700
        .Height = 225
        .BorderStyle = 0
    End With
    
    lblCurrentPos.Top = 0
    
    lngSuggestedPos = Me.Width - lblCurrentPos.Width - 50
    If fraLegend.Visible Then
        lngMinimumPos = fraLegend.Left + fraLegend.Width + 50
    Else
        lngMinimumPos = 120
    End If
    If lngSuggestedPos < lngMinimumPos Then
        lngSuggestedPos = lngMinimumPos
    End If
    
    lblCurrentPos.Left = lngSuggestedPos
    
    With fraPlot
        .Left = 50
        .Width = Me.Width - 300
        
        PlotHeight = Me.Height - 700
        lngSuggestedPos = 0
        If fraLegend.Visible Or lblCurrentPos.Visible Then
            lngSuggestedPos = lngSuggestedPos + fraLegend.Height
            PlotHeight = PlotHeight - fraLegend.Height + 45
        Else
            lngSuggestedPos = 50
        End If
        .Top = lngSuggestedPos
        
        If PlotHeight < 1 Then PlotHeight = 1
        .Height = PlotHeight
        
'''        dynMSPlot.Left = .Left
'''        dynMSPlot.top = .top
'''        dynMSPlot.Width = .Width
'''        dynMSPlot.Height = .Height

    End With


End Sub

Private Sub RedrawZoomBox()
    Dim TheseCoords As usrRect
    
    TheseCoords = ZoomBoxCoords
            
    With TheseCoords
        If .x1 > .x2 Then
            SwapValues .x1, .x2
        End If
        If .y1 > .y2 Then
            SwapValues .y1, .y2
        End If
        
    End With
        
    ' When the box size gets large enough, turn on blnDrawingZoomBox
    If IsZoomBoxLargeEnough(TheseCoords) Then
        blnDrawingZoomBox = True
    End If
    
    With shpZoomBox
        .Left = TheseCoords.x1
        .Top = TheseCoords.y1
        .Width = TheseCoords.x2 - TheseCoords.x1
        .Height = TheseCoords.y2 - TheseCoords.y1
    End With
    
End Sub

Public Sub RefreshPlot(Optional blnZoomOutCompletely As Boolean = True)
    If blnZoomOutCompletely Then
        ZoomOut blnZoomOutCompletely
    Else
        UpdatePlot False
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 32 Then
        SetCursorMode True
        SetCursorIcon
    ElseIf KeyCode = vbKeyEscape Then
        mnuClose_Click
    ElseIf KeyCode = 90 And (Shift And vbCtrlMask) Then
        ' Ctrl+Z
        ZoomOut False
    End If
    
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 32 Then
        SetCursorMode False
        SetCursorIcon
    End If
End Sub

Private Sub Form_Load()
    SizeAndCenterWindow Me, cWindowBottomRight, 8000, 7000, True
    
    tmrUpdatePlot.Interval = 100
    
    InitializePlotOptions
    
    shpZoomBox.Visible = False
    
    EnableDisableMenuCheckmarks
    EnableDisableZoomMenus False
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ResetMousePointer Button
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
        HidePlotForm
    End If
End Sub

Private Sub Form_Resize()
    PositionControls
    If Me.WindowState <> vbMinimized Then
        TickCountToUpdateAt = TicksElapsedSinceStart + 1
        blnResizingWindow = True
    End If
End Sub

Public Function GetPlotType() As Integer
    GetPlotType = PlotOptions.PlotTypeCode
End Function

Public Function GetResolution() As Long
    GetResolution = PlotOptions.GaussianConversion.ResolvingPower
End Function

Public Function GetResolutionXVal() As Double
    GetResolutionXVal = PlotOptions.GaussianConversion.XValueOfSpecification
End Function

Public Function GetGridlinesStatus(blnXAxis As Boolean) As Boolean
    If blnXAxis Then
        GetGridlinesStatus = PlotOptions.XAxis.ShowGridLinesMajor
    Else
        GetGridlinesStatus = PlotOptions.YAxis.ShowGridLinesMajor
    End If
End Function

Public Function GetAxisTickCount(blnXAxis As Boolean) As Integer
    If blnXAxis Then
        GetAxisTickCount = PlotOptions.XAxis.MajorTicksToShow
    Else
        GetAxisTickCount = PlotOptions.YAxis.MajorTicksToShow
    End If
End Function

Public Function GetGaussianQualityFactor() As Integer
    GetGaussianQualityFactor = PlotOptions.GaussianConversion.QualityFactor
End Function

Public Function GetPlotApproximationFactor() As Integer
    GetPlotApproximationFactor = PlotOptions.ApproximationFactor
End Function
Public Function GetAutoscaleYStatus() As Boolean
    GetAutoscaleYStatus = PlotOptions.AutoScaleY
End Function

Public Function GetPositionDisplayStatus() As Boolean
    GetPositionDisplayStatus = blnUpdatePosition
End Function

Private Sub fraPlot_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    Dim boolClickInsideBox As Boolean
    boolClickInsideBox = IsClickInBox(x, y, ZoomBoxCoords)
    
    If Not shpZoomBox.Visible Or Not boolClickInsideBox Then
        InitializeZoomOrMove Button, x, y
    Else
        ' The Zoom Box is visible
        ' The click is also handled in Sub fraPlot_MouseUp since it is more customary to handle clicks with _MouseUp events
        If blnZoomBoxDrawn = True Then
            RespondZoomModeClick Button, x, y
        End If
    End If
End Sub

Private Sub fraPlot_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 0 Then
        SetCursorIcon x, y
    End If
    
    ResizeZoomBox Button, x, y
    
    blnUpdatePosition = True
    CurrentPosX = x
    CurrentPosY = y

End Sub

Private Sub fraPlot_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim intHistoryIndex As Integer, intDataSetIndex As Integer
    
    If mnuCursorModeZoom.Checked Then
        ' Zooming, the box is shown at the size the user chose
        ' Handle the click differently depending on the location and the button
        
        If blnZoomBoxDrawn = True Then
            ' Do nothing
        ElseIf blnDrawingZoomBox Then
            blnZoomBoxDrawn = True
            EnableDisableZoomMenus IsZoomBoxLargeEnough(ZoomBoxCoords)
        End If
    Else
        ' Moving plot
        ' Turn off blnSlidingGraph
        blnSlidingGraph = False
        
        ' Set TickCountToUpdateAt back to 0 so no more udpates occur
        TickCountToUpdateAt = 0
        
        ' Update the plot one more time so new view gets saved to history
        UpdatePlot True
        
        ' For some reason the history still has two copies of the most recent view
        ' Remove one of them
        For intHistoryIndex = 2 To PLOT_RANGE_HISTORY_COUNT - 1
            For intDataSetIndex = 1 To intDataSetLoadedCount - 1
                PlotRangeHistory(intHistoryIndex, intDataSetIndex) = PlotRangeHistory(intHistoryIndex + 1, intDataSetIndex)
            Next intDataSetIndex
        Next intHistoryIndex
    End If
    
End Sub

Private Sub mnuApproximationFactor_Click()
    SetPlotApproximationFactor
End Sub

Private Sub mnuAutoScaleYAxis_Click()
    SetAutoscaleY Not mnuAutoScaleYAxis.Checked
End Sub

Private Sub mnuClose_Click()
    HidePlotForm
End Sub

Private Sub mnuCursorModeMove_Click()
    SetCursorMode True
End Sub

Private Sub mnuCursorModeZoom_Click()
    SetCursorMode False
End Sub

Private Sub mnuExportData_Click()
    ExportData
End Sub

Private Sub mnuFile_Click()
    ' EnableDisable mnuExportData
    
    EnableDisableExportDataMenu
End Sub

Private Sub mnuFixMinimumYAtZero_Click()
    SetFixMinimumYAtZero Not mnuFixMinimumYAtZero.Checked
End Sub

Private Sub mnuGaussianQuality_Click()
    SetGaussianQuality
End Sub

Private Sub mnuGridLinesXAxis_Click()
    SetGridlinesXAxis Not mnuGridLinesXAxis.Checked
End Sub

Private Sub mnuGridLinesYAxis_Click()
    SetGridlinesYAxis Not mnuGridLinesYAxis.Checked
End Sub

Private Sub mnuOptions_Click()
    EnableDisableMenuCheckmarks
End Sub

Private Sub mnuPlotTypeGaussianPeaks_Click()
    SetPlotType True
End Sub

Private Sub mnuPlotTypeSticksToZero_Click()
    SetPlotType False
End Sub

Private Sub mnuResetToDefaults_Click()
    Dim eResponse As VbMsgBoxResult
    
    eResponse = MsgBox(LookupMessage(1140), vbYesNoCancel + vbDefaultButton2 + vbQuestion, LookupMessage(1145))
    
    If eResponse = vbYes Then
        InitializePlotOptions
        SetPlotType CBool(PlotOptions.PlotTypeCode), False
        ZoomOut True
    End If
End Sub

Private Sub mnuSetRangeX_Click()
    SetNewRange True, False
End Sub

Private Sub mnuSetRangeY_Click()
    SetNewRange False, True
End Sub

Private Sub mnuSetResolution_Click()
    SetResolution
End Sub

Private Sub mnuShowCurrentPosition_Click()
    SetShowCursorPosition Not mnuShowCurrentPosition.Checked
End Sub

Private Sub mnuShowLegend_Click()
    SetLegendVisibility Not mnuShowLegend.Checked
End Sub

'''Private Sub mnuShowDynaPlot_Click()
'''    mnuShowDynaPlot.Checked = Not mnuShowDynaPlot.Checked
'''
'''    fraPlot.Visible = Not mnuShowDynaPlot.Checked
'''    dynMSPlot.Visible = mnuShowDynaPlot.Checked
'''End Sub
'''
'''Private Sub dynMSPlot_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'''    If Button <> vbLeftButton Then
'''        ' Zoom out to previous zoom
'''        'dynMSPlot.Scale.ZoomOut
'''        dynMSPlot.Action.CallActionFunction acZOOMOUT
'''    End If
'''End Sub
'''
'''Private Sub dynMSPlot_Zoomed(ByVal XLeft As Double, ByVal XRight As Double, ByVal YBottom As Double, ByVal YTop As Double)
'''
'''    With dynMSPlot.Scale
'''        .XLeft = XLeft
'''        .XRight = XRight
'''        .YBottom = 0
'''        .YTop = 100
'''    End With
'''
'''End Sub

Private Sub mnuTicksXAxis_Click()
    AxisTickCountUpdateByUser True, PlotOptions.XAxis.MajorTicksToShow
End Sub

Private Sub mnuTicksYAxis_Click()
    AxisTickCountUpdateByUser False, PlotOptions.YAxis.MajorTicksToShow
End Sub

Private Sub mnuZoomIn_Click()
    HideZoomBox vbLeftButton, True
End Sub

Private Sub mnuZoomInHorizontal_Click()
    ZoomInHorizontal
End Sub

Private Sub mnuZoomInVertical_Click()
    ZoomInVertical
End Sub

Private Sub mnuZoomOut_Click()
    ZoomShrink False, False
End Sub

Private Sub mnuZoomOutFullScale_Click()
    ZoomOut True
End Sub

Private Sub mnuZoomOutHorizontal_Click()
    ZoomShrink True, False
End Sub

Private Sub mnuZoomOutToPrevious_Click()
    ZoomOut False
End Sub

Private Sub mnuZoomOutVertical_Click()
    ZoomShrink False, True
End Sub

Private Sub tmrUpdatePlot_Timer()
    ' Note: the internal for the timer is 100 msec
    
    TicksElapsedSinceStart = TicksElapsedSinceStart + 1
    If TickCountToUpdateAt > 0 Then
        UpdatePlotWhenIdle
    End If
    
    If blnUpdatePosition Then UpdateCurrentPos
End Sub

