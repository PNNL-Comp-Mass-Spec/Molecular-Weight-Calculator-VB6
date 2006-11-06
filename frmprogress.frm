VERSION 5.00
Begin VB.Form frmProgress 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Progress"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   7350
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   7350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Tag             =   "14700"
   Begin VB.Frame fraControls 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   735
      Left            =   3840
      TabIndex        =   5
      Top             =   1560
      Width           =   2415
      Begin VB.CommandButton cmdPause 
         Caption         =   "Click to Pause"
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Tag             =   "14710"
         Top             =   0
         Width           =   1815
      End
      Begin VB.Label lblPressEscape 
         Alignment       =   2  'Center
         Caption         =   "(Press Escape to abort)"
         Height          =   255
         Left            =   0
         TabIndex        =   7
         Tag             =   "14730"
         Top             =   480
         Width           =   2655
      End
   End
   Begin VB.Shape pbarSubProgressBox 
      Height          =   255
      Left            =   120
      Top             =   960
      Width           =   2775
   End
   Begin VB.Shape pbarSubProgress 
      BackColor       =   &H80000002&
      FillColor       =   &H80000002&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   120
      Top             =   960
      Width           =   1215
   End
   Begin VB.Shape pbarProgressBox 
      Height          =   255
      Left            =   120
      Top             =   120
      Width           =   2775
   End
   Begin VB.Shape pbarProgress 
      BackColor       =   &H80000002&
      FillColor       =   &H80000002&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   120
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblSubtaskProgress 
      Alignment       =   2  'Center
      Caption         =   "Current Task Progress"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Tag             =   "14730"
      Top             =   1350
      Width           =   2655
   End
   Begin VB.Label lblOverallProgress 
      Alignment       =   2  'Center
      Caption         =   "Overall Progress"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Tag             =   "14730"
      Top             =   510
      Width           =   2655
   End
   Begin VB.Label lblCurrentSubTask 
      Height          =   735
      Left            =   3120
      TabIndex        =   4
      Top             =   720
      Width           =   4095
   End
   Begin VB.Label lblTimeStats 
      Caption         =   "Elapsed/remaining time"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1800
      Width           =   3015
   End
   Begin VB.Label lblCurrentTask 
      Caption         =   "Current task"
      Height          =   495
      Left            =   3120
      TabIndex        =   3
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Written by Matthew Monroe for use in applications
' First written in Chapel Hill, NC in roughly 2000
'
' Last Modified:    July 24, 2003

Private ePauseStatus As ppProgressPauseConstants

Private lngProgressMin As Long
Private lngProgressMax As Long
Private dblLatestProgressBarValue As Double

Private lngSubTaskProgressMin As Long
Private lngSubTaskProgressMax As Long
Private dblLatestSubTaskProgressBarValue As Double

Private mblnWorking As Boolean
Private mTextMinutesElapsedRemaining As String
Private mTextPreparingToPause As String
Private mTextResuming As String
Private mTextClickToPause As String
Private mTextPaused As String
Private mTextPressEscapeToAbort As String
Private mLastQueryUnloadTickCount As Long

Private Enum ppProgressPauseConstants
    ppUnpaused = 0
    ppRequestPause
    ppPaused
    ppRequestUnpause
End Enum

Private Sub CheckForPauseUnpause()
    
    Select Case ePauseStatus
    Case ppRequestPause
        cmdPause.Caption = mTextPaused
        ePauseStatus = ppPaused
        Me.MousePointer = vbNormal
        Do
            Sleep 100
            DoEvents
        Loop While ePauseStatus = ppPaused
        ePauseStatus = ppUnpaused
        cmdPause.Caption = mTextClickToPause
        Me.MousePointer = vbHourglass
    Case ppRequestUnpause
        ePauseStatus = ppUnpaused
        cmdPause.Caption = mTextClickToPause
    Case Else
        ' Nothing to pause or unpause
    End Select

End Sub

Public Function GetElapsedTime() As Single
    ' Examines lblTimeStats to find the last recorded elapsed time
    
    Dim intColonLoc As Integer
    
    intColonLoc = InStr(lblTimeStats, ":")
    
    If intColonLoc > 0 Then
        GetElapsedTime = Val(Left(lblTimeStats, intColonLoc - 1))
    End If
    
End Function

Public Function GetProgressBarValue(ByRef lngMinimum As Long, ByRef lngMaximum As Long) As Double
    ' Returns the current value of the progress bar, plus the Min and Max, returning them ByRef
    lngMinimum = lngProgressMin
    lngMaximum = lngProgressMax
    
    GetProgressBarValue = dblLatestProgressBarValue
End Function

Public Function GetSubTaskProgressBarValue(ByRef lngMinimum As Long, ByRef lngMaximum As Long) As Double
    ' Returns the current value of the subtask progress bar, plus the Min and Max, returning them ByRef
    lngMinimum = lngSubTaskProgressMin
    lngMaximum = lngSubTaskProgressMax
    
    GetSubTaskProgressBarValue = dblLatestSubTaskProgressBarValue
End Function

Public Sub HideForm(Optional blnResetKeyPressAbortProcess As Boolean = True)
    
    If blnResetKeyPressAbortProcess Then KeyPressAbortProcess = 0
    
    frmProgress.MousePointer = vbNormal
    
    mblnWorking = False
    
    ' The following On Error is necessary in case a modal window is displayed
    ' Also necessary since a call to .Hide when a form is already hidden generates an error
    On Error Resume Next
    frmProgress.Hide
    
End Sub

Public Sub InitializeSubtask(CurrentSubTask As String, SubTaskProgressBarMinNew As Long, SubTaskProgressBarMaxNew As Long)
    lngSubTaskProgressMin = SubTaskProgressBarMinNew
    lngSubTaskProgressMax = SubTaskProgressBarMaxNew
    
    If lngSubTaskProgressMin > lngSubTaskProgressMax Then
        ' Swap them
        SwapValues lngSubTaskProgressMin, lngSubTaskProgressMax
    End If
    
    If lngSubTaskProgressMin < 0 Then lngSubTaskProgressMin = 0
    If lngSubTaskProgressMin > lngSubTaskProgressMax Then lngSubTaskProgressMax = lngSubTaskProgressMin + 1
    If lngSubTaskProgressMax < 1 Then lngSubTaskProgressMax = 1
    
On Error GoTo InitializeSubtaskErrorHandler

    UpdateCurrentSubTask CurrentSubTask
    
    UpdateSubtaskProgressBar lngSubTaskProgressMin
    
    mblnWorking = True

    Exit Sub
    
InitializeSubtaskErrorHandler:
    MsgBox "An error occurred while initializing the sub-progress bar: " & vbCrLf & Err.Description, vbInformation + vbOKOnly, "Error"
    Resume Next
    
End Sub

Public Sub InitializeForm(CurrentTask As String, OverallProgressBarMinNew As Long, OverallProgressBarMaxNew As Long, Optional blnShowTimeStats As Boolean = False, Optional blnShowSubTaskProgress As Boolean = False, Optional blnShowPauseButton As Boolean = True, Optional frmObjOwnerForm As Form)
    lngProgressMin = OverallProgressBarMinNew
    lngProgressMax = OverallProgressBarMaxNew
    
    If lngProgressMin > lngProgressMax Then
        ' Swap them
        SwapValues lngProgressMin, lngProgressMax
    End If
    
    If lngProgressMin < 0 Then lngProgressMin = 0
    If lngProgressMin > lngProgressMax Then lngProgressMax = lngProgressMin + 1
    If lngProgressMax < 1 Then lngProgressMax = 1
    
    lngSubTaskProgressMin = 0
    lngSubTaskProgressMax = 1
    pbarSubProgress.Width = 0
    
On Error GoTo InitializeFormErrorHandler

    UpdateCurrentTask CurrentTask
    lblCurrentSubTask = ""
    
    lblTimeStats.Visible = blnShowTimeStats
    
    pbarSubProgress.Visible = blnShowSubTaskProgress
    pbarSubProgressBox.Visible = blnShowSubTaskProgress
    lblSubtaskProgress.Visible = blnShowSubTaskProgress
    
    If blnShowSubTaskProgress Then
        lblTimeStats.Top = 1800
        fraControls.Top = 1650
        fraControls.Left = 3840
        Me.Height = 2800
    Else
        lblTimeStats.Top = 800
        fraControls.Top = 1190
        fraControls.Left = 240
        Me.Height = 2350
    End If
    
    cmdPause.Visible = blnShowPauseButton
    If ppPaused Then
        ePauseStatus = ppRequestUnpause
        CheckForPauseUnpause
    End If
    
    UpdateProgressBar lngProgressMin, True
    
    KeyPressAbortProcess = 0
    mLastQueryUnloadTickCount = 0
    mblnWorking = True
    
    If frmObjOwnerForm Is Nothing Then
        frmProgress.Show vbModeless
    Else
        frmProgress.Show vbModeless, frmObjOwnerForm
    End If
    
    frmProgress.MousePointer = vbHourglass

	Exit Sub

InitializeFormErrorHandler:
    If Err.Number = 401 Then
        ' Tried to show frmProgress when a modal form is shown; this isn't allowed
        ' Probably cannot use frmProgress in the calling routine
        Debug.Assert False
        Resume Next
    Else
        MsgBox "An error occurred while initializing the progress bar: " & vbCrLf & Err.Description, vbInformation + vbOKOnly, "Error"
        Resume Next
    End If
    
End Sub

Public Sub MoveToBottomCenter()
    SizeAndCenterWindow Me, cWindowBottomCenter, -1, -1, False
End Sub

Public Function TaskInProgress() As Boolean
    ' Returns True if the Progress form is currently displayed
    ' Returns False otherwise

    TaskInProgress = mblnWorking
End Function

Public Sub ToggleAlwaysOnTop(blnStayOnTop As Boolean)
    Static blnCurrentlyOnTop As Boolean
    
    If blnCurrentlyOnTop = blnStayOnTop Then Exit Sub
    
    Me.ScaleMode = vbTwips
    
    WindowStayOnTop Me.hWnd, blnStayOnTop, Me.ScaleX(Me.Left, vbTwips, vbPixels), Me.ScaleY(Me.Top, vbTwips, vbPixels), Me.ScaleX(Me.Width, vbTwips, vbPixels), Me.ScaleY(Me.Height, vbTwips, vbPixels)
    
    blnCurrentlyOnTop = blnStayOnTop
End Sub

Public Sub UpdateProgressBar(ByVal NewValue As Double, Optional ResetStartTime As Boolean = False)
    
    Static StartTime As Double
    Static StopTime As Double
    
    Dim MinutesElapsed As Double, MinutesTotal As Double, MinutesRemaining As Double
    Dim dblRatioCompleted As Double
    
    If ResetStartTime Then
        StartTime = Now()
    End If
    
    dblRatioCompleted = SetProgressBarValue(NewValue, False)
    
    StopTime = Now()
    MinutesElapsed = (StopTime - StartTime) * 1440
    If dblRatioCompleted <> 0 Then
        MinutesTotal = MinutesElapsed / dblRatioCompleted
    Else
        MinutesTotal = 0
    End If
    MinutesRemaining = MinutesTotal - MinutesElapsed
    lblTimeStats = Format(MinutesElapsed, "0.00") & " : " & Format(MinutesRemaining, "0.00 ") & mTextMinutesElapsedRemaining
    
    CheckForPauseUnpause
    
    DoEvents
    
End Sub

Public Sub SetStandardCaptionText(Optional ByVal strMinutesElapsedRemaining As String = "min. elapsed/remaining", Optional ByVal strPreparingToPause As String = "Preparing to Pause", Optional ByVal strResuming As String = "Resuming", Optional ByVal strClickToPause As String = "Click to Pause", Optional ByVal strPaused As String = "Paused", Optional ByVal strPressEscapeToAbort As String = "(Press Escape to abort)")
    mTextMinutesElapsedRemaining = strMinutesElapsedRemaining
    mTextPreparingToPause = strPreparingToPause
    mTextResuming = strResuming
    mTextClickToPause = strClickToPause
    mTextPaused = strPaused
    mTextPressEscapeToAbort = strPressEscapeToAbort
    
    lblPressEscape.Caption = mTextPressEscapeToAbort
End Sub

Private Function SetProgressBarValue(ByRef dblNewValue As Double, Optional blnIncrementalUpdate As Boolean = False) As Double
    ' Updates the value of the Progress bar based on dblNewValue, adjusting if necessary
    ' Returns the % completed ratio
    ' If blnIncrementalUpdate is true, then does not update dblLatestProgressBarValue
    
    Dim dblRatioCompleted As Double
    
    If dblNewValue < lngProgressMin Then dblNewValue = lngProgressMin
    If dblNewValue > lngProgressMax Then dblNewValue = lngProgressMax
    
    If lngProgressMax > lngProgressMin Then
        dblRatioCompleted = (dblNewValue - lngProgressMin) / (lngProgressMax - lngProgressMin)
    Else
        dblRatioCompleted = 0
    End If
    If dblRatioCompleted < 0 Then dblRatioCompleted = 0
    If dblRatioCompleted > 1 Then dblRatioCompleted = 1
    
On Error GoTo ExitSetProgressBar
    
    pbarProgress.Width = pbarProgressBox.Width * dblRatioCompleted
    
    If Not blnIncrementalUpdate Then
        dblLatestProgressBarValue = dblNewValue
    End If

    SetProgressBarValue = dblRatioCompleted
    
ExitSetProgressBar:

End Function

Public Sub UpdateSubtaskProgressBar(ByVal dblNewValue As Double, Optional ByVal blnFractionallyIncreaseOverallProgressValue As Boolean = True)
    Dim dblPartialIncrementToAdd As Double, dblNewTotalProgressValue As Double
    Dim dblRatioCompleted As Double
    Dim lngSubtaskProgressBarLength As Long
    
    If dblNewValue < lngSubTaskProgressMin Then dblNewValue = lngSubTaskProgressMin
    If dblNewValue > lngSubTaskProgressMax Then dblNewValue = lngSubTaskProgressMax
    
    dblLatestSubTaskProgressBarValue = dblNewValue
    
    If lngSubTaskProgressMax > 0 Then
        dblRatioCompleted = (dblNewValue - lngSubTaskProgressMin) / lngSubTaskProgressMax
    Else
        dblRatioCompleted = 0
    End If
    If dblRatioCompleted < 0 Then dblRatioCompleted = 0
    If dblRatioCompleted > 1 Then dblRatioCompleted = 1
    
    On Error GoTo ExitUpdateSubProgressBarFunction
    
    pbarSubProgress.Width = pbarSubProgressBox.Width * dblRatioCompleted
    
    lngSubtaskProgressBarLength = lngSubTaskProgressMax - lngSubTaskProgressMin
    If lngSubtaskProgressBarLength > 0 And blnFractionallyIncreaseOverallProgressValue Then
        dblPartialIncrementToAdd = (dblNewValue - lngSubTaskProgressMin) / CDbl(lngSubtaskProgressBarLength)
        
        dblNewTotalProgressValue = dblLatestProgressBarValue + dblPartialIncrementToAdd
        If dblNewTotalProgressValue > dblLatestProgressBarValue Then
            SetProgressBarValue dblNewTotalProgressValue, True
        End If
    End If
     
    CheckForPauseUnpause
    
    DoEvents

ExitUpdateSubProgressBarFunction:

End Sub

Public Sub UpdateCurrentTask(strNewTask As String)
    lblCurrentTask = strNewTask
    
    CheckForPauseUnpause
    
    DoEvents
End Sub

Public Sub UpdateCurrentSubTask(strNewSubTask As String)
    lblCurrentSubTask = strNewSubTask
    
    CheckForPauseUnpause
    
    DoEvents
End Sub

Public Sub WasteTime(Optional Milliseconds As Integer = 250)
    ' Wait the specified number of milliseconds
    
    ' Use of the Sleep API call is more efficient than using a VB timer since it results in 0% processor usage
    Sleep Milliseconds
    
End Sub

Private Sub cmdPause_Click()
    Select Case ePauseStatus
    Case ppUnpaused
        ePauseStatus = ppRequestPause
        cmdPause.Caption = mTextPreparingToPause
        DoEvents
    Case ppPaused
        ePauseStatus = ppRequestUnpause
        cmdPause.Caption = mTextResuming
        DoEvents
    Case Else
        ' Ignore click
    End Select
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        KeyPressAbortProcess = 2
        KeyCode = 0
        Shift = 0
        
        ' Move the form to the bottom center of the screen to avoid having the msgbox popup under the form
        MoveToBottomCenter
    End If
End Sub

Private Sub Form_Load()
    
    ' Put window in exact center of screen
    SizeAndCenterWindow Me, cWindowLowerThird, 7450, 2800, False

    SetStandardCaptionText
    
    mblnWorking = False
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim eResponse As VbMsgBoxResult
    Dim lngCurrentTickCount As Long

    If UnloadMode = vbFormControlMenu Then
        ' If at least 20 seconds has elapsed since the last time the user tried
        '   to close the form, then query whether they really want to do this
        lngCurrentTickCount = GetTickCount()
        If Abs(lngCurrentTickCount - mLastQueryUnloadTickCount) >= 20000 Then
            ' Move the form to the bottom center of the screen to avoid having the msgbox popup under the form
            Me.MoveToBottomCenter
            
            eResponse = MsgBox("It appears that a task is currently in progress.  Do you really want to close the progress window?  If yes, this will cancel the currently running task.", vbQuestion + vbYesNo + vbDefaultButton2, "Cancel Task")
            If eResponse = vbYes Then
                ' Set KeyPressAbortProcess to 2 so that the program will cancel the task gracefully, and (hopefully) hide the progress window
                KeyPressAbortProcess = 2
            Else
                ' User said no; set lngCurrentTickCount to 0 to guarantee MsgBox will reoccur if user clicks again
                lngCurrentTickCount = 0
            End If
    
            Cancel = 1
        End If
        mLastQueryUnloadTickCount = lngCurrentTickCount
    End If
End Sub
