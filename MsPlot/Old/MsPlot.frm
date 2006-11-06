VERSION 5.00
Begin VB.Form frmMsPlot 
   Caption         =   "Plot"
   ClientHeight    =   3495
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3495
   ScaleWidth      =   6000
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraPlot 
      BorderStyle     =   0  'None
      Height          =   2985
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4785
      Begin VB.Timer tmrUpdatePlot 
         Interval        =   100
         Left            =   4320
         Top             =   0
      End
      Begin VB.Frame fraOptions 
         Caption         =   "Options (hidden)"
         Height          =   1275
         Left            =   1560
         TabIndex        =   4
         Top             =   1320
         Visible         =   0   'False
         Width           =   2175
         Begin VB.TextBox txtXAxisTickCount 
            Height          =   285
            Left            =   1440
            TabIndex        =   7
            Text            =   "5"
            Top             =   120
            Width           =   615
         End
         Begin VB.TextBox txtYAxisTickCount 
            Height          =   285
            Left            =   1440
            TabIndex        =   6
            Text            =   "5"
            Top             =   480
            Width           =   615
         End
         Begin VB.ComboBox cboLabelsToShow 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   840
            Width           =   615
         End
         Begin VB.Label lblXAxisTickCount 
            Caption         =   "# of x axis ticks"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label lblYAxisTickCount 
            Caption         =   "# of y axis ticks"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label lblLabelsToShow 
            Caption         =   "# of ions to label:"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   840
            Width           =   1215
         End
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
      Begin VB.Label lblCurrentPos 
         Caption         =   "Loc: 0,0"
         Height          =   195
         Left            =   1680
         TabIndex        =   3
         Top             =   0
         Width           =   1995
      End
      Begin VB.Label lblTick 
         BackStyle       =   0  'Transparent
         Caption         =   "123.45"
         Height          =   200
         Index           =   1
         Left            =   720
         TabIndex        =   2
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
      Begin VB.Shape shpZoomBox 
         BorderColor     =   &H000000FF&
         BorderStyle     =   4  'Dash-Dot
         Height          =   855
         Left            =   240
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label lblPlotIntensity 
         BackStyle       =   0  'Transparent
         Caption         =   "123.43"
         Height          =   195
         Index           =   1
         Left            =   2160
         TabIndex        =   1
         Top             =   360
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Line linGridline 
         BorderColor     =   &H00FF0000&
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
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuInitializeData 
         Caption         =   "&Initialize with Dummy Data"
      End
      Begin VB.Menu mnuInitializeStickDataLots 
         Caption         =   "Initialize with MS Stick Data (lots)"
      End
      Begin VB.Menu mnuInitializeStickData 
         Caption         =   "Initialize with MS Stick Data (just a few)"
      End
      Begin VB.Menu mnuExportData 
         Caption         =   "&Export Data"
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
         Begin VB.Menu mnuPlotTypeLinesBetweenPoints 
            Caption         =   "&Lines Between Points"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuSetResolution 
         Caption         =   "Set Effective Resolution"
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
         Caption         =   "&Ticks to show (approx.)"
         Begin VB.Menu mnuTicksXAxis 
            Caption         =   "&X Axis..."
         End
         Begin VB.Menu mnuTicksYAxis 
            Caption         =   "&Y Axis..."
         End
      End
      Begin VB.Menu mnuPeaksToLabel 
         Caption         =   "Peaks To &Label..."
         Enabled         =   0   'False
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
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuFixMinimumYAtZero 
         Caption         =   "&Fix mimimum Y at zero"
      End
      Begin VB.Menu mnuOptionsSepBar5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuZoomOutToPrevious 
         Caption         =   "&Zoom Out to Previous   "
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
         Caption         =   "Show Current &Position"
         Checked         =   -1  'True
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

Private ZoomBoxCoords As usrRect
Private PlotOptions As usrPlotDataOptions

Private boolSlidingGraph As Boolean, PlotRangeAtMoveStart As usrPlotRange
Private boolResizingWindow As Boolean, boolDrawingZoomBox As Boolean, boolZoomBoxDrawn As Boolean

Private LoadedXYData As usrXYDataSet        ' The data to plot
Private InitialStickData As usrXYDataSet    ' If the user submits Stick Data (discrete data points) and requests that the sticks be converted to a
                                            '   Gaussian representation, then the original, unmodified data is stored here

Private dblPlotRangeStretchVal As Double

Public TicksElapsedSinceStart As Long      ' Actually increments 10 times per second rather than 1000 per second since tmrPlot.Interval = 100
Private TickCountToUpdateAt As Long

Private boolUpdatePosition As Boolean, CurrentPosX As Double, CurrentPosY As Double

Const PLOT_RANGE_HISTORY_COUNT = 20
Private PlotRangeHistory(PLOT_RANGE_HISTORY_COUNT) As usrPlotRange        ' Keeps track of the last 5 plot ranges displayed to allow for undoing

Private Sub EnableDisableZoomMenus(boolEnableMenus As Boolean)

    mnuZoomOptions.Visible = boolEnableMenus

End Sub

Private Sub EnableDisableMenuCheckmarks()
    
    With PlotOptions
        mnuPlotTypeSticksToZero.Checked = (.PlotTypeCode = 0)
        mnuPlotTypeLinesBetweenPoints.Checked = (.PlotTypeCode = 1)
        mnuGridLinesXAxis.Checked = .XAxis.ShowGridLinesMajor
        mnuGridLinesYAxis.Checked = .YAxis.ShowGridLinesMajor
        mnuAutoScaleYAxis.Checked = .AutoScaleY
        mnuFixMinimumYAtZero.Checked = .FixYAxisMinimumAtZero
    End With
    
End Sub

Public Sub ExportData()
    Dim lngIndex As Long, strFilepath As String
    
    On Error GoTo WriteProblem
    
    strFilepath = "d:\DataOut.csv"
    Open strFilepath For Output As #1
    For lngIndex = 1 To LoadedXYData.XYDataListCount
        With LoadedXYData.XYDataList(lngIndex)
            Print #1, .XVal & "," & .YVal
        End With
    Next lngIndex
    Close

    MsgBox "Data written to file " & strFilepath

    Exit Sub
    
WriteProblem:
    MsgBox "Error writing data to file " & strFilepath
    
End Sub

Private Function FixUpCoordinates(TheseCoords As usrRect) As usrRect
    Dim FixedCoords As usrRect

    FixedCoords = TheseCoords
    
    With FixedCoords
        If .x1 > .x2 Then
            SwapValues .x1, .x2
        End If
        If .y1 > .y2 Then
            SwapValues .y1, .y2
        End If
    End With
    
    FixUpCoordinates = FixedCoords
    
End Function

Private Sub HideZoomBox(Button As Integer, boolPerformZoom As Boolean)
    
    EnableDisableZoomMenus False
    boolZoomBoxDrawn = False
    boolDrawingZoomBox = False
    
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
            boolZoomBoxDrawn = False
            EnableDisableZoomMenus False
            RedrawZoomBox
        Else
            ' Moving (dragging) plot
            ' Will still update the zoom/move box to keep track of how far dragged
            ' but will not draw the box
            shpZoomBox.Visible = False
            boolDrawingZoomBox = False
            boolSlidingGraph = True
            PlotRangeAtMoveStart = PlotRangeHistory(1)
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
        If boolZoomBoxDrawn Then
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

Public Sub SetFixMinimumAtZero(boolEnable As Boolean)
    PlotOptions.FixYAxisMinimumAtZero = boolEnable
    SetPlotOptions True, True
    EnableDisableMenuCheckmarks
End Sub

Public Sub SetGridlinesXAxis(boolEnable As Boolean)
    PlotOptions.XAxis.ShowGridLinesMajor = boolEnable
    SetPlotOptions True, False
    EnableDisableMenuCheckmarks
End Sub

Public Sub SetGridlinesYAxis(boolEnable As Boolean)
    PlotOptions.YAxis.ShowGridLinesMajor = boolEnable
    SetPlotOptions True, False
    EnableDisableMenuCheckmarks
End Sub

Public Sub SetNewRange(boolIsXAxis As Boolean, boolPromptUserForValues As Boolean, Optional dblNewStartVal As Double = 0, Optional dblNewEndVal As Double = 0)
    Dim dblStartVal As Double, dblEndVal As Double, dblLowerLimit As Double, dblUpperLimit As Double
    Dim dblDefaultSeparationValue As Double
    Dim strFormatString As String
    
    If boolIsXAxis Then
        dblStartVal = PlotOptions.Scaling.x.ValStart.Val - PlotOptions.Scaling.x.ValNegativeValueCorrectionOffset
        dblEndVal = PlotOptions.Scaling.x.ValEnd.Val - PlotOptions.Scaling.x.ValNegativeValueCorrectionOffset
        dblLowerLimit = PlotOptions.DataLimits.x.ValStart.Val
        dblUpperLimit = PlotOptions.DataLimits.x.ValEnd.Val
    Else
        dblStartVal = PlotOptions.Scaling.y.ValStart.Val - PlotOptions.Scaling.y.ValNegativeValueCorrectionOffset
        dblEndVal = PlotOptions.Scaling.y.ValEnd.Val - PlotOptions.Scaling.y.ValNegativeValueCorrectionOffset
        dblLowerLimit = PlotOptions.DataLimits.y.ValStart.Val
        dblUpperLimit = PlotOptions.DataLimits.y.ValEnd.Val
    End If
    
    If dblStartVal = 0 And dblEndVal = 0 Then
        dblStartVal = dblLowerLimit
        dblEndVal = dblUpperLimit
    End If
    
    If boolPromptUserForValues Or (dblNewStartVal = 0 And dblNewEndVal = 0) Then
            
        With PlotRangeHistory(1)
            If boolIsXAxis Then
                strFormatString = ConstructFormatString(Abs(.x.ValEnd.Val - .x.ValStart.Val) / 100)
            Else
                strFormatString = ConstructFormatString(Abs(.y.ValEnd.Val - .y.ValStart.Val) / 100)
            End If
        End With
            
        dblStartVal = Format(dblStartVal, strFormatString)
        dblEndVal = Format(dblEndVal, strFormatString)
        
        With frmSetValue
            .Caption = "Set Range"
            .lblStartVal.Caption = "Start Val"
            .txtStartVal = dblStartVal
            .lblEndVal.Caption = "End Val"
            .txtEndVal = dblEndVal
            
            ' Round dblDefaultSeparationValue to nearest 1, 2, or 5 (or multiple of 10 thereof)
            dblDefaultSeparationValue = RoundToMultipleOf10((dblEndVal - dblStartVal) / 10)
            
            .SetLimits True, dblLowerLimit, dblUpperLimit, dblDefaultSeparationValue
            
            If boolIsXAxis Then
                .Caption = "Set X Axis Range"
            Else
                .Caption = "Set Y Axis Range"
            End If
        
            .Show vbModal
        End With
        
        If UCase(frmSetValue.Tag) <> "OK" Then Exit Sub
        
        ' Set New Range
        With frmSetValue
            If IsNumeric(.txtStartVal) Then dblNewStartVal = CDbl(.txtStartVal)
            If IsNumeric(.txtEndVal) Then dblNewEndVal = CDbl(.txtEndVal)
        End With
    End If
    
    If boolIsXAxis Then
        PlotOptions.Scaling.x.ValStart.Val = dblNewStartVal
        PlotOptions.Scaling.x.ValEnd.Val = dblNewEndVal
    Else
        PlotOptions.Scaling.y.ValStart.Val = dblNewStartVal
        PlotOptions.Scaling.y.ValEnd.Val = dblNewEndVal
    End If
    
    UpdatePlot True
End Sub

Public Sub SetXYData(NewXYData As usrXYDataSet, boolTreatDataAsDiscretePoints As Boolean, Optional boolConvertStickDataToGaussianRepresentation As Boolean = False)
    
    ' Reset .DataLimits
    With PlotOptions.DataLimits
        .x.ValStart.Val = 0
        .y.ValEnd.Val = 0
        .y.ValStart.Val = 0
        .y.ValEnd.Val = 0
    End With

    ' Initialize Plot options
    InitializePlotOptions
    
    If boolTreatDataAsDiscretePoints Then
        InitialStickData = NewXYData
        ' Load New Data into LoadedXYData
        SetPlotOptions False, False
        SetPlotType boolConvertStickDataToGaussianRepresentation, False
    Else
        InitialStickData.XYDataListCount = 0
        LoadedXYData = NewXYData
        SetPlotType True, False             ' Note: Automatically loads data into LoadedXYData
    End If
    
    PlotOptions.Scaling = PlotOptions.DataLimits
    
    ' Reset the boolLongOperationsRequired bit
    PlotOptions.boolLongOperationsRequired = False
    
    ' Erase all data in PlotRangeHistory()
    ' Since this array is dimensioned using the const PLOT_RANGE_HISTORY_COUNT, it
    '  does not need to be re-dimensioned after erasing
    Erase PlotRangeHistory()
    
End Sub

Public Sub SetPeaksToLabel(Optional intPeaksToLabel As Integer = -1)
    Dim strResponse As String, intNewLabelCount As Integer
    
    If intPeaksToLabel < 0 Then
        strResponse = InputBox("Please enter the number of peaks (sticks) to label by decreasing intensity:", "Peaks to Label", cboLabelsToShow.ListIndex)
    Else
        strResponse = Trim(Str(intPeaksToLabel))
    End If
    
    If IsNumeric(strResponse) Then
        intNewLabelCount = CInt(strResponse)
        If intNewLabelCount < 0 Or intNewLabelCount >= cboLabelsToShow.ListCount Then
            intNewLabelCount = 3
        End If
        If intNewLabelCount < cboLabelsToShow.ListCount Then
            cboLabelsToShow.ListIndex = intNewLabelCount
        End If
    End If
    
    SetPlotOptions True, False
End Sub
Private Sub InitializePlotOptions()
    txtXAxisTickCount = "5"
    txtYAxisTickCount = "5"
    cboLabelsToShow.ListIndex = 0
    
    With PlotOptions
        .XAxis.Show = True
        .XAxis.MinorTickMinimumPixelSep = 100
        .XAxis.ShowMinorTicks = True
        
        .YAxis.Show = True
        .YAxis.MinorTickMinimumPixelSep = 100
        .YAxis.ShowMinorTicks = True
            
        .GaussianConversion.ResolvingPower = 5000
        .GaussianConversion.XValueOfSpecification = 500
        .HighlightColor = vbRed
    End With

End Sub

Private Sub SetPlotOptions(Optional boolUpdatePlot As Boolean = True, Optional boolUpdateHistory As Boolean = True)
    With PlotOptions
        .XAxis.MajorTicksToShow = txtXAxisTickCount
        
        .YAxis.MajorTicksToShow = txtYAxisTickCount
        
        .LabelsToShow = Val(cboLabelsToShow)
        .ShowDataPointLabels = (.LabelsToShow > 0)
    End With
    
    If boolUpdatePlot Then UpdatePlot boolUpdateHistory
    
End Sub

Public Sub SetPlotType(boolLinesBetweenPoints As Boolean, Optional boolUpdatePlot As Boolean = True)
    
    If boolLinesBetweenPoints Then
        PlotOptions.PlotTypeCode = 1
    Else
        PlotOptions.PlotTypeCode = 0
    End If
    mnuPeaksToLabel.Enabled = Not boolLinesBetweenPoints
    
    If InitialStickData.XYDataListCount > 0 Then
        ' Stick data is present; need to take action
        If boolLinesBetweenPoints Then
            LoadedXYData = ConvertStickDataToGaussian(Me, InitialStickData, PlotOptions)
        Else
            LoadedXYData = InitialStickData
        End If
    End If
    
    SetPlotOptions boolUpdatePlot, False
    EnableDisableMenuCheckmarks
    
End Sub

Public Sub SetResolution(Optional lngNewResolvingPower As Long = -1, Optional dblNewXValResLocation As Double = 500)
    Dim strResponse As String
    
    If lngNewResolvingPower < 1 Then
        With frmSetValue
            .Caption = "Resolving Power Specifications"
            .lblStartVal.Caption = "Resolving Power"
            .txtStartVal = PlotOptions.GaussianConversion.ResolvingPower
            .lblEndVal.Caption = "X Value of Specification"
            .txtEndVal = PlotOptions.GaussianConversion.XValueOfSpecification
            
            .SetLimits False
            
            .Show vbModal
        End With
        
        If UCase(frmSetValue.Tag) <> "OK" Then Exit Sub
        
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
    PlotOptions.GaussianConversion.ResolvingPower = lngNewResolvingPower
    
    PlotOptions.GaussianConversion.XValueOfSpecification = dblNewXValResLocation
    
    SetPlotOptions False, False
    SetPlotType mnuPlotTypeLinesBetweenPoints.Checked, True

End Sub

Public Sub SetShowCursorPosition(boolEnable As Boolean)
    mnuShowCurrentPosition.Checked = boolEnable
    EnableDisableMenuCheckmarks
End Sub

Private Sub TickCountUpdateByUser(txtThisTextBox As TextBox, Optional strAxisLetter As String = "X")
    Dim strResponse As String, intNewTickCount As Integer
    
    strAxisLetter = UCase(Trim(strAxisLetter))
    
    If Len(strAxisLetter) = 0 Then strAxisLetter = "X"
    
    strResponse = InputBox("Please enter the approximate number of ticks to show on the " & strAxisLetter & " axis (2 to 30):", strAxisLetter & " Axis Ticks", txtThisTextBox)
    
    If IsNumeric(strResponse) Then
        intNewTickCount = CInt(strResponse)
        If intNewTickCount < 2 Or intNewTickCount > 30 Then
            intNewTickCount = 5
        End If
        txtThisTextBox = Trim(Str(intNewTickCount))
    End If
    
    SetPlotOptions True, False

End Sub

Private Sub UpdateCurrentPos()
    Dim XValue As Double, YValue As Double
    Dim strFormatStringX As String, strFormatStringY As String
    
    If mnuShowCurrentPosition.Checked Then
        With PlotRangeHistory(1)
            XValue = XYPosToValue(CLng(CurrentPosX), .x)
            YValue = XYPosToValue(CLng(CurrentPosY), .y)
            strFormatStringX = ConstructFormatString(Abs(.x.ValEnd.Val - .x.ValStart.Val) / 100)
            strFormatStringY = ConstructFormatString(Abs(.y.ValEnd.Val - .y.ValStart.Val) / 100)
        End With
        
        lblCurrentPos = "Loc: " & Format(XValue, strFormatStringX) & ", " & Format(YValue, strFormatStringY)
    End If
    
    boolUpdatePosition = False
End Sub

Private Sub UpdatePlot(Optional boolUpdateHistory As Boolean = True)
    Dim MostRecentPlotRange As usrPlotRange, intIndex As Integer
    
    With PlotOptions
        If .PlotTypeCode = 0 Then
            .PlotTop = 500
        Else
            .PlotTop = 250
        End If
        .PlotLeft = 300
        .PlotWidth = fraPlot.Width - .PlotLeft - 250
        .PlotHeight = fraPlot.Height - .PlotTop - 700
    End With
    
    MostRecentPlotRange = PlotRangeHistory(1)
    
    ' Hide the plot so it gets updated faster
    fraPlot.Visible = False
    
    ' Perform the actual update
    DrawPlot Me, PlotOptions, LoadedXYData.XYDataList, LoadedXYData.XYDataListCount, MostRecentPlotRange
        
    ' Show the plot
    fraPlot.Visible = True

    If boolUpdateHistory Then
'        If PlotRangeHistory(1).X.ValEnd.Val <> PlotRangeHistory(2).X.ValEnd.Val Or _
            PlotRangeHistory(1).X.ValStart.Val <> PlotRangeHistory(2).X.ValStart.Val Or _
            PlotRangeHistory(1).Y.ValEnd.Val <> PlotRangeHistory(2).Y.ValEnd.Val Or _
            PlotRangeHistory(1).Y.ValStart.Val <> PlotRangeHistory(2).Y.ValStart.Val Then
            ' Update the plot range history
            For intIndex = PLOT_RANGE_HISTORY_COUNT To 2 Step -1
                PlotRangeHistory(intIndex) = PlotRangeHistory(intIndex - 1)
            Next intIndex
'        End If
    End If
    
    If Me.MousePointer = vbHourglass Then Me.MousePointer = vbDefault
    
    PlotRangeHistory(1) = MostRecentPlotRange
End Sub

Private Sub UpdatePlotWhenIdle()
    Dim boolUpdateHistory As Boolean
    
    If TickCountToUpdateAt < TicksElapsedSinceStart Then
        ' Update the plot, but don't update the history if sliding
        If boolSlidingGraph Or boolResizingWindow Then
            boolUpdateHistory = False
        Else
            boolUpdateHistory = True
        End If
        UpdatePlot boolUpdateHistory
        
        TickCountToUpdateAt = 0
        boolResizingWindow = False
    End If
    
End Sub

Private Sub ZoomInHorizontal()
    ' Zoom in along the horizontal axis but
    ' Do not change the vertical range
    
    FixUpCoordinates ZoomBoxCoords
    With ZoomBoxCoords
        .y1 = PlotRangeHistory(1).y.ValEnd.Pos
        .y2 = PlotRangeHistory(1).y.ValStart.Pos
    End With
     
    HideZoomBox vbLeftButton, True

End Sub

Private Sub ZoomInVertical()
    ' Zoom in along the horizontal axis but
    ' Do not change the vertical range
    
    FixUpCoordinates ZoomBoxCoords
    With ZoomBoxCoords
        .x1 = PlotRangeHistory(1).x.ValStart.Pos
        .x2 = PlotRangeHistory(1).x.ValEnd.Pos
    End With
     
    HideZoomBox vbLeftButton, True

End Sub

Private Sub ZoomOut(ByVal boolZoomOutCompletely As Boolean)
    Dim intIndex As Integer
    
    If Not boolZoomOutCompletely Then
        ' See if any previous PlotRange data exists in the history
        ' If not then set boolZoomOutCompletely to True
        With PlotRangeHistory(2)
            If .x.ValStart.Pos = 0 And .x.ValEnd.Pos = 0 And .x.ValStart.Pos = 0 And .x.ValEnd.Pos = 0 Then
                ' Most recent saved zoom range is all zeroes -- not usable
               boolZoomOutCompletely = True
            End If
        End With
    End If
    
    If boolZoomOutCompletely Then
        ' Call SetPlotOptions to make sure all options are up to date
        SetPlotOptions False
        
        With PlotOptions
            ' Override the AutoScaleY option and turn on ZoomOutFull
            .ZoomOutFull = True
            .AutoScaleY = True
            
            ' Initialize .Scaling.x.ValStart.Val and .Scaling.y.ValStart.Val if necessary
            If LoadedXYData.XYDataListCount > 0 Then
                .Scaling.x.ValStart.Val = LoadedXYData.XYDataList(1).XVal
                .Scaling.x.ValEnd.Val = LoadedXYData.XYDataList(LoadedXYData.XYDataListCount).XVal
            End If
            
            If .PlotTypeCode = 0 Then
                ' Displaying a sticks to zero plot and zoomed out full
                ' Need to stretch the limits of the plot by 2% of the total range
                dblPlotRangeStretchVal = (.Scaling.x.ValEnd.Val - .Scaling.x.ValStart.Val) * 0.05
                .Scaling.x.ValEnd.Val = .Scaling.x.ValEnd.Val + dblPlotRangeStretchVal
                .Scaling.x.ValStart.Val = .Scaling.x.ValStart.Val - dblPlotRangeStretchVal
            End If

        End With
        

        ' Update theplot
        UpdatePlot
        
        ' Call SetPlotOptions again in case .AutoScaleY should be false
        SetPlotOptions False
    Else
        ' Zoom to previous range
        
        PlotOptions.Scaling = PlotRangeHistory(2)
        
        UpdatePlot False
        
        ' Update the plot range history
        For intIndex = 2 To PLOT_RANGE_HISTORY_COUNT - 1
            PlotRangeHistory(intIndex) = PlotRangeHistory(intIndex + 1)
        Next intIndex
        
    End If

End Sub

Private Sub ZoomShrink(boolFixHorizontal As Boolean, boolFixVertical As Boolean)
    ' Zoom out, but not completely
    
    Dim lngViewRangePosX As Long, lngViewRangePosY As Long
    Dim lngBoxSizeX As Double, lngBoxSizeY As Double
    Dim lngPosCorrectionFactorX As Long, lngPosCorrectionFactorY As Long
    Dim TheseCoords As usrRect
    
    lngViewRangePosX = PlotRangeHistory(1).x.ValEnd.Pos - PlotRangeHistory(1).x.ValStart.Pos
    lngViewRangePosY = PlotRangeHistory(1).y.ValEnd.Pos - PlotRangeHistory(1).y.ValStart.Pos
    
    If lngViewRangePosX = 0 Or lngViewRangePosY = 0 Then Exit Sub
    
    TheseCoords = ZoomBoxCoords
    FixUpCoordinates TheseCoords
    With TheseCoords
        lngBoxSizeX = Abs(.x2 - .x1)
        lngBoxSizeY = Abs(.y2 - .y1)
        If boolFixVertical Then
            .x1 = PlotRangeHistory(1).x.ValStart.Pos
            .x2 = PlotRangeHistory(1).x.ValEnd.Pos
        Else
            If lngBoxSizeX > 0 Then
                lngPosCorrectionFactorX = (CLng(lngViewRangePosX * CDbl(lngViewRangePosX) / CDbl(lngBoxSizeX))) / 2
                .x1 = .x1 - lngPosCorrectionFactorX
                .x2 = .x2 + lngPosCorrectionFactorX
            End If
        End If
        If boolFixHorizontal Then
            .y1 = PlotRangeHistory(1).y.ValEnd.Pos
            .y2 = PlotRangeHistory(1).y.ValStart.Pos
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
           y >= .y1 And y <= .y2 Then
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
    Dim DeltaXVal As Double, DeltaYVal As Double
    Dim MinimumXVal As Double, MaximumXVal As Double, MaximumRange As Double
    
    TheseCoords = ZoomBoxCoords
            
    With PlotRangeAtMoveStart
        DeltaXVal = XYPosToValue(TheseCoords.x2, .x) - XYPosToValue(TheseCoords.x1, .x)
        DeltaYVal = XYPosToValue(TheseCoords.y2, .y) - XYPosToValue(TheseCoords.y1, .y)
    End With
    
    PlotOptions.ZoomOutFull = False
    With PlotOptions
        MaximumRange = .DataLimits.x.ValEnd.Val - .DataLimits.x.ValStart.Val
        .Scaling.x.ValStart.Val = PlotRangeAtMoveStart.x.ValStart.Val - DeltaXVal
        MinimumXVal = .DataLimits.x.ValStart.Val - MaximumRange / 10
        If .Scaling.x.ValStart.Val < MinimumXVal Then
            .Scaling.x.ValStart.Val = MinimumXVal
            .Scaling.x.ValEnd.Val = .Scaling.x.ValStart.Val + (PlotRangeAtMoveStart.x.ValEnd.Val - PlotRangeAtMoveStart.x.ValStart.Val)
        Else
            .Scaling.x.ValEnd.Val = PlotRangeAtMoveStart.x.ValEnd.Val - DeltaXVal
        End If
        
        MaximumXVal = .DataLimits.x.ValEnd.Val + MaximumRange / 10
        If .Scaling.x.ValEnd.Val > MaximumXVal Then
            .Scaling.x.ValEnd.Val = MaximumXVal
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
    Dim TheseCoords As usrRect, boolAbortZoom As Boolean
    Dim PlotRangeSaved As usrPlotRange
    
    ' Use the numbers stored in PlotRangeSaved to update the PlotOptions with the desired zoom range
    PlotRangeSaved = PlotRangeHistory(1)    ' The most recent plot range
    
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
    
    With fraOptions
        .Left = 50
        .Top = Me.Height - .Height - 700
        .Visible = False
    End With
    
    With fraPlot
        .Top = 50
        .Left = 50
        .Width = Me.Width - 250
        If fraOptions.Visible Then
            PlotHeight = Me.Height - fraOptions.Height - 600
        Else
            PlotHeight = Me.Height - 600
        End If
        
        If PlotHeight < 1 Then PlotHeight = 1
        .Height = PlotHeight
        
        lblCurrentPos.Top = 0
        lblCurrentPos.Left = fraPlot.Width - lblCurrentPos.Width - 100
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
        
    ' When the box size gets large enough, turn on boolDrawingZoomBox
    If IsZoomBoxLargeEnough(TheseCoords) Then
        boolDrawingZoomBox = True
    End If
    
    With shpZoomBox
        .Left = TheseCoords.x1
        .Top = TheseCoords.y1
        .Width = TheseCoords.x2 - TheseCoords.x1
        .Height = TheseCoords.y2 - TheseCoords.y1
    End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 32 Then
        SetCursorMode True
        SetCursorIcon
    ElseIf KeyCode = vbKeyEscape Then
        mnuClose_Click
    ElseIf KeyCode = 90 And (Shift Or vbCtrlMask) Then
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
    SizeAndCenterWindow Me, cWindowExactCenter, 8000, 7000, True
    
    Dim x As Integer
    
    With cboLabelsToShow
        For x = 0 To 20
            .AddItem Trim(Str(x))
        Next x
        .ListIndex = 3
    End With

    InitializePlotOptions
    
    shpZoomBox.Visible = False
    
    mnuCursorMode.Caption = mnuCursorMode.Caption & vbTab & "Space Enables Move"
    mnuZoomIn.Caption = mnuZoomIn.Caption & vbTab & "Left Click"
    mnuZoomOutToPrevious.Caption = mnuZoomOutToPrevious.Caption & vbTab & "Ctrl+Z or Right Click"
    
    EnableDisableZoomMenus False
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ResetMousePointer Button
End Sub

Private Sub Form_Resize()
    PositionControls
    If Me.WindowState <> vbMinimized Then
        TickCountToUpdateAt = TicksElapsedSinceStart + 1
        boolResizingWindow = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim objThisForm As Form
    
    For Each objThisForm In Forms
        If objThisForm.Name <> Me.Name Then
            Unload objThisForm
        End If
    Next
End Sub

Private Sub fraOptions_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ResetMousePointer Button
End Sub

Private Sub fraPlot_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    Dim boolClickInsideBox As Boolean
    boolClickInsideBox = IsClickInBox(x, y, ZoomBoxCoords)
    
    If Not shpZoomBox.Visible Or Not boolClickInsideBox Then
        InitializeZoomOrMove Button, x, y
    Else
        ' The Zoom Box is visible
        ' The click is also handled in Sub fraPlot_MouseUp since it is more customary to handle clicks with _MouseUp events
        If boolZoomBoxDrawn = True Then
            RespondZoomModeClick Button, x, y
        End If
    End If
End Sub

Private Sub fraPlot_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 0 Then
        SetCursorIcon x, y
    End If
    
    ResizeZoomBox Button, x, y
    
    boolUpdatePosition = True
    CurrentPosX = x
    CurrentPosY = y

End Sub

Private Sub fraPlot_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim intIndex As Integer
    
    If mnuCursorModeZoom.Checked Then
        ' Zooming, the box is shown at the size the user chose
        ' Handle the click differently depending on the location and the button
        
        If boolZoomBoxDrawn = True Then
            ' Do nothing
        ElseIf boolDrawingZoomBox Then
            boolZoomBoxDrawn = True
            EnableDisableZoomMenus IsZoomBoxLargeEnough(ZoomBoxCoords)
        End If
    Else
        ' Moving plot
        ' Turn off boolSlidingGraph
        boolSlidingGraph = False
        
        ' Set TickCountToUpdateAt back to 0 so no more udpates occur
        TickCountToUpdateAt = 0
        
        ' Update the plot one more time so new view gets saved to history
        UpdatePlot True
        
        ' For some reason the history still has two copies of the most recent view
        ' Remove one of them
        For intIndex = 2 To PLOT_RANGE_HISTORY_COUNT - 1
            PlotRangeHistory(intIndex) = PlotRangeHistory(intIndex + 1)
        Next intIndex
    End If
    
End Sub


Private Sub mnuAutoScaleYAxis_Click()
    SetAutoscaleY Not mnuAutoScaleYAxis.Checked
End Sub

Private Sub mnuClose_Click()
    Unload Me
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

Private Sub mnuFixMinimumYAtZero_Click()
    SetFixMinimumAtZero Not mnuFixMinimumYAtZero.Checked
End Sub

Private Sub mnuGridLinesXAxis_Click()
    SetGridlinesXAxis Not mnuGridLinesXAxis.Checked
End Sub

Private Sub mnuGridLinesYAxis_Click()
    SetGridlinesYAxis Not mnuGridLinesYAxis.Checked
End Sub

Private Sub mnuInitializeData_Click()
    InitializeDummyData 0
End Sub

Private Sub mnuInitializeStickData_Click()
    InitializeDummyData 1
End Sub

Private Sub mnuInitializeStickDataLots_Click()
    InitializeDummyData 2

End Sub

Private Sub mnuOptions_Click()
    EnableDisableMenuCheckmarks
End Sub

Private Sub mnuPeaksToLabel_Click()
    SetPeaksToLabel
End Sub

Private Sub mnuPlotTypeLinesBetweenPoints_Click()
    SetPlotType True
End Sub

Private Sub mnuPlotTypeSticksToZero_Click()
    SetPlotType False
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

Private Sub mnuTicksXAxis_Click()
    TickCountUpdateByUser txtXAxisTickCount, "X"
End Sub

Private Sub mnuTicksYAxis_Click()
    TickCountUpdateByUser txtYAxisTickCount, "Y"
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
    ' Note: the internal for the timer is 250 msec
    
    TicksElapsedSinceStart = TicksElapsedSinceStart + 1
    If TickCountToUpdateAt > 0 Then
        UpdatePlotWhenIdle
    End If
    
    If boolUpdatePosition Then UpdateCurrentPos
End Sub

Private Sub txtXAxisTickCount_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtXAxisTickCount, KeyAscii, True, False
End Sub

Private Sub txtXAxisTickCount_Validate(Cancel As Boolean)
    ValidateTextboxValueDbl txtXAxisTickCount, 2, 50, 10
    SetPlotOptions True, False
End Sub

Private Sub txtYAxisTickCount_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtYAxisTickCount, KeyAscii, True, False
End Sub

Private Sub txtYAxisTickCount_Validate(Cancel As Boolean)
    ValidateTextboxValueDbl txtYAxisTickCount, 2, 50, 10
    SetPlotOptions True, False
End Sub
