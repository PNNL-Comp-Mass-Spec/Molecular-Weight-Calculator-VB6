VERSION 5.00
Begin VB.Form frmProgress 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Progress"
   ClientHeight    =   2580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3105
   Icon            =   "Progress.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   3105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Tag             =   "14700"
   Begin VB.CommandButton cmdPause 
      Caption         =   "Click to Pause"
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Tag             =   "14710"
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Timer tmrDelayTimer 
      Enabled         =   0   'False
      Left            =   2400
      Top             =   1680
   End
   Begin VB.Shape pbarProgress 
      BackColor       =   &H80000002&
      FillColor       =   &H80000002&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   120
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Shape pbarBox 
      Height          =   255
      Left            =   120
      Top             =   1200
      Width           =   2775
   End
   Begin VB.Label lblCurrentSubTask 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   2655
   End
   Begin VB.Label lblTimeStats 
      Caption         =   "Elapsed/remaining time"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1515
      Width           =   3015
   End
   Begin VB.Label lblPressEscape 
      Alignment       =   2  'Center
      Caption         =   "(Press Escape to abort)"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Tag             =   "14730"
      Top             =   2280
      Width           =   2655
   End
   Begin VB.Label lblCurrentTask 
      Caption         =   "Current task"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private DelayTimerIntervalCount As Integer
Private intPauseStatus As Integer

Private lngProgressMin As Long
Private lngProgressMax As Long

Private Const cUnpaused = 0
Private Const cRequestPause = 1
Private Const cPaused = 2
Private Const cRequestUnpause = 3

Private Sub CheckForPauseUnpause()
    
    Select Case intPauseStatus
    Case cRequestPause
        cmdPause.Caption = LookupLanguageCaption(14720, "Paused")
        intPauseStatus = cPaused
        Me.MousePointer = vbNormal
        Do
            WasteTime 100
            DoEvents
        Loop While intPauseStatus = cPaused
        intPauseStatus = cUnpaused
        cmdPause.Caption = LookupLanguageCaption(14710, "Click to Pause")
        Me.MousePointer = vbHourglass
    Case cRequestUnpause
        intPauseStatus = cUnpaused
    Case Else
        ' Nothing to pause or unpause
    End Select

End Sub

Public Sub InitializeForm(CurrentTask As String, ByVal ProgressBarMinNew As Long, ByVal ProgressBarMaxNew As Long, Optional boolShowTimeStats As Boolean = False)
    If ProgressBarMinNew < 0 Then ProgressBarMinNew = 0
    If ProgressBarMinNew > ProgressBarMaxNew Then ProgressBarMaxNew = ProgressBarMinNew + 1
    
    lblCurrentTask.Caption = CurrentTask
    lblCurrentSubTask = ""
    
    lngProgressMin = ProgressBarMinNew
    lngProgressMax = ProgressBarMaxNew
    
    If lngProgressMin < 0 Then lngProgressMin = 0
    If lngProgressMax <= lngProgressMin Then lngProgressMax = lngProgressMin + 1
    
    lblTimeStats.Visible = boolShowTimeStats
    
    UpdateProgressBar ProgressBarMinNew, True
    
    KeyPressAbortProcess = 0
    
    frmProgress.Show
    frmProgress.MousePointer = vbHourglass
    
End Sub

Public Sub UpdateProgressBar(ByVal NewValue As Long, Optional ResetStartTime As Boolean = False)
    
    Static StartTime As Double
    Static StopTime As Double
    
    Dim MinutesElapsed As Currency, MinutesTotal As Currency, MinutesRemaining As Currency
    Dim dblRatioCompleted As Double, lngNewWidth As Long
    
    If ResetStartTime Then
        StartTime = Now()
    End If
    
    If NewValue < lngProgressMin Then NewValue = lngProgressMin
    If NewValue > lngProgressMax Then NewValue = lngProgressMax
    
    If lngProgressMax > 0 Then
        dblRatioCompleted = (NewValue - lngProgressMin) / lngProgressMax
    Else
        dblRatioCompleted = 0
    End If
    If dblRatioCompleted < 0 Then dblRatioCompleted = 0
    If dblRatioCompleted > 1 Then dblRatioCompleted = 1
    pbarProgress.Width = pbarBox.Width * dblRatioCompleted
    
    On Error GoTo ExitUpdateProgressBarFunction
    
    StopTime = Now()
    MinutesElapsed = (StopTime - StartTime) * 1440
    If dblRatioCompleted <> 0 Then
        MinutesTotal = MinutesElapsed / dblRatioCompleted
    Else
        MinutesTotal = 0
    End If
    MinutesRemaining = MinutesTotal - MinutesElapsed
    lblTimeStats = Format(MinutesElapsed, "0.00") & " : " & Format(MinutesRemaining, "0.00 ") & LookupLanguageCaption(14740, "min. elapsed/remaining")
    
    CheckForPauseUnpause
    
    DoEvents
    
ExitUpdateProgressBarFunction:
    
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
    
    Const Default_Interval = 10
    
    Dim dblStopCount As Double
    
    With frmProgress
        DelayTimerIntervalCount = 0
        
        .tmrDelayTimer.Interval = Default_Interval
        .tmrDelayTimer.Enabled = True
    
        dblStopCount = Milliseconds / Default_Interval
        Do While DelayTimerIntervalCount < dblStopCount
            DoEvents
        Loop
        .tmrDelayTimer.Enabled = False
        
    End With
    
End Sub

Private Sub cmdPause_Click()
    Select Case intPauseStatus
    Case cUnpaused
        intPauseStatus = cRequestPause
        cmdPause.Caption = LookupLanguageCaption(14715, "Preparing to Pause")
        DoEvents
    Case cPaused
        intPauseStatus = cRequestUnpause
        cmdPause.Caption = LookupLanguageCaption(14725, "Resuming")
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
    End If
End Sub

Private Sub Form_Load()
    
    ' Put window in exact center of screen
    SizeAndCenterWindow Me, cWindowExactCenter, 3200, 3000, False

End Sub


Private Sub tmrDelayTimer_Timer()
    Dim IntervalCount As Integer
    
    DelayTimerIntervalCount = DelayTimerIntervalCount + 1
    
    If DelayTimerIntervalCount > 32767 Then DelayTimerIntervalCount = -32767
End Sub


