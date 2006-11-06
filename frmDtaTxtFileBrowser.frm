VERSION 5.00
Begin VB.Form frmDtaTxtFileBrowser 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "_Dta.Txt File Browser"
   ClientHeight    =   1320
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   10515
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1320
   ScaleWidth      =   10515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Tag             =   "16000"
   Begin VB.CommandButton cmdJumpToScan 
      Cancel          =   -1  'True
      Caption         =   "&Jump to scan"
      Height          =   360
      Left            =   7560
      TabIndex        =   10
      Tag             =   "16100"
      ToolTipText     =   "Shortcut is Ctrl+J"
      Top             =   720
      Width           =   1275
   End
   Begin VB.CheckBox chkWindowStayOnTop 
      Caption         =   "&Keep Window On Top"
      Height          =   495
      Left            =   9000
      TabIndex        =   11
      Tag             =   "16110"
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Cl&ose"
      Height          =   360
      Left            =   7560
      TabIndex        =   9
      Tag             =   "4000"
      Top             =   240
      Width           =   1275
   End
   Begin VB.Frame fraDTATextOptions 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7215
      Begin VB.TextBox txtParentIonCharge 
         Height          =   285
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   660
         Width           =   1455
      End
      Begin VB.TextBox txtParentIonMass 
         Height          =   285
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   240
         Width           =   1455
      End
      Begin VB.ComboBox cboDTAScanNumber 
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblParentIonCharge 
         Caption         =   "Parent Ion Charge"
         Height          =   255
         Left            =   3600
         TabIndex        =   7
         Tag             =   "16080"
         Top             =   660
         Width           =   1575
      End
      Begin VB.Label lblParentIonMH 
         Caption         =   "Parent Ion MH+"
         Height          =   255
         Left            =   3600
         TabIndex        =   3
         Tag             =   "16070"
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblScanNumberEndLabel 
         Caption         =   "Scan Number End"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Tag             =   "16060"
         Top             =   660
         Width           =   1575
      End
      Begin VB.Label lblScanNumberEnd 
         Caption         =   "End Scan Number"
         Height          =   255
         Left            =   1920
         TabIndex        =   6
         Top             =   660
         Width           =   1455
      End
      Begin VB.Label lblDTAScanNumber 
         Caption         =   "Scan Number Start"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Tag             =   "16050"
         Top             =   240
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmDtaTxtFileBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type udtSpectraType
    ScanNumberStart As Long
    ScanNumberEnd As Long
    ParentIon As Double
    ParentIonCharge As Integer
    DataCount As Long
    XVals() As Double
    YVals() As Double
End Type

Private Type udtDtaTxtDataType
    Initialized As Boolean
    SpectraCount As Long
    Spectra() As udtSpectraType
    DtaTxtFilePath As String
End Type

Private mDtaTxtData As udtDtaTxtDataType

Public Function GetDataInitializedState() As Boolean
    GetDataInitializedState = mDtaTxtData.Initialized
End Function

Public Sub JumpToScanNumber(Optional ByRef lngScanNumber As Long = 0)
    ' If lngScanNumber is greater than 0, then jumps directly to the given scan number
    ' Otherwise, shows an Input box to allow the user to type a scan number
    ' The scan number entered by the user is returned ByRef in lngScanNumber
    
    Dim strFirstScan As String, strLastScan As String
    Dim strScanNumber As String
    Dim lngIndex As Long
    Dim blnMatched As Boolean
    
    If cboDTAScanNumber.ListCount < 1 Then
        If lngScanNumber = 0 Then MsgBox "The scan list is empty; no scans to jump to.", vbInformation + vbOKOnly, "No Scans in Memory"
        Exit Sub
    End If
    
    If lngScanNumber <= 0 Then
        strFirstScan = cboDTAScanNumber.List(0)
        strLastScan = cboDTAScanNumber.List(cboDTAScanNumber.ListCount - 1)
        
        strScanNumber = InputBox("Enter the scan number to jump to (" & strFirstScan & " to " & strLastScan & "): ", "Jump to Scan", cboDTAScanNumber)
        If IsNumeric(strScanNumber) Then
            lngScanNumber = CLngSafe(strScanNumber)
            
            With cboDTAScanNumber
                If CLngSafe(.List(0)) > lngScanNumber Then
                    .ListIndex = 0
                Else
                    For lngIndex = 0 To .ListCount - 1
                        If CLngSafe(.List(lngIndex)) >= lngScanNumber Then
                            .ListIndex = lngIndex
                            blnMatched = True
                            Exit For
                        End If
                    Next lngIndex
                    
                    If Not blnMatched Then .ListIndex = .ListCount - 1
                End If
            End With
        
        End If
    End If
End Sub

Public Sub ReadDtaTxtFile(strFilePath As String)

    Dim InFileNum As Integer
    Dim lngDtaTxtSpectraDimCount As Long
    
    Dim lngTotalByteCount As Long, lngTotalBytesRead As Long
    Dim lngCurrentLineNumber As Long
    Dim intParseCount As Integer, intIndex As Integer
    Dim strParsedVals() As String
    Dim lngDataDimCount As Long
    Dim lngIndex As Long
    
    Dim strLineIn As String
    Dim blnSkipNextRead As Boolean
    Dim lngScanNumberStart As Long
    Dim lngScanNumberEnd As Long
    Dim intChargeFromHeader As Integer
    Dim dblParentIonMass As Double
    Dim intChargeFromParentIon As Integer
    Dim blnSkipThisScan As Boolean
    
    Const SCAN_DELIMETER = "="

On Error GoTo ReadDtaTxtFileErrorHandler

    ' Initialize mDtaTxtData
    With mDtaTxtData
        .Initialized = True
        .SpectraCount = 0
        .DtaTxtFilePath = strFilePath
        lngDtaTxtSpectraDimCount = 1000
        ReDim .Spectra(lngDtaTxtSpectraDimCount)
    End With

    ' Make sure the file exists
    If Not FileExists(strFilePath) Then
        mDtaTxtData.Initialized = False
        Exit Sub
    End If

    InFileNum = FreeFile()
    Open strFilePath For Input As #InFileNum
    lngTotalByteCount = FileLen(strFilePath)

    frmProgress.InitializeForm "Reading _Dta.Txt file", 0, lngTotalByteCount, True, False, True

    lngTotalBytesRead = 0
    lngCurrentLineNumber = 0
    Do While Not EOF(InFileNum)
        If blnSkipNextRead Then
            blnSkipNextRead = False
        Else
            Line Input #InFileNum, strLineIn
            strLineIn = Trim(strLineIn)
        End If
        
        lngCurrentLineNumber = lngCurrentLineNumber + 1
        lngTotalBytesRead = lngTotalBytesRead + Len(strLineIn) + 2      ' Add 2 bytes to account for CrLf at end of line

        If lngCurrentLineNumber Mod 1000 = 0 Then
            ' Only update the progress bar every 250 lines
            frmProgress.UpdateProgressBar lngTotalBytesRead
            If KeyPressAbortProcess > 1 Then Exit Do
        End If

        If Left(strLineIn, 1) = SCAN_DELIMETER Then
            ' Header line found
            ' Determine scan range and assigned charge
            lngScanNumberStart = 0
            lngScanNumberEnd = 0
            intChargeFromHeader = 0
            dblParentIonMass = 0
            intChargeFromParentIon = 0
            
            intParseCount = ParseString(strLineIn, strParsedVals(), 15, "." & Chr(34) & " ", "", False, False, False)
            
            ' Find the enry in strParsedVals() that contains dta
            For intIndex = intParseCount - 1 To 0 Step -1
                If LCase(strParsedVals(intIndex)) = "dta" Then
                    If intIndex >= 3 Then
                        Debug.Assert intIndex = 6
                        
                        lngScanNumberStart = CLngSafe(strParsedVals(intIndex - 3))
                        lngScanNumberEnd = CLngSafe(strParsedVals(intIndex - 2))
                        intChargeFromHeader = CLngSafe(strParsedVals(intIndex - 1))
                    End If
                    Exit For
                End If
            Next intIndex
                    
            ' Read next line to get Parent ion (as M+H) and charge
            Line Input #InFileNum, strLineIn
            lngCurrentLineNumber = lngCurrentLineNumber + 1
            lngTotalBytesRead = lngTotalBytesRead + Len(strLineIn) + 2      ' Add 2 bytes to account for CrLf at end of line

            intParseCount = ParseString(strLineIn, strParsedVals(), 3, " ", "", True, True, False)

            If intParseCount >= 2 Then
                dblParentIonMass = CDblSafe(strParsedVals(0))
                intChargeFromParentIon = CIntSafe(strParsedVals(1))
                If intChargeFromParentIon <> intChargeFromHeader Then
                    intChargeFromHeader = intChargeFromParentIon
                End If
            End If

            With mDtaTxtData
                ' Check if this scan has the same values for lngScanNumberStart and lngScanNumberEnd as the previous one
                blnSkipThisScan = False
                If .SpectraCount > 0 Then
                    With .Spectra(.SpectraCount - 1)
                        If .ScanNumberStart = lngScanNumberStart And .ScanNumberEnd = lngScanNumberEnd Then
                            blnSkipThisScan = True
                        End If
                    End With
                End If
                
                If blnSkipThisScan Then
                    ' Skip all of the lines for this scan by continuing to read data until a line
                    '  a line is found that doesn't start with a number
                    blnSkipNextRead = False
                    Do While Not EOF(InFileNum)
                        Line Input #InFileNum, strLineIn
                        strLineIn = Trim(strLineIn)
                        If Not IsNumeric(Left(strLineIn, 1)) Then
                            blnSkipNextRead = True
                            Exit Do
                        Else
                            lngCurrentLineNumber = lngCurrentLineNumber + 1
                            lngTotalBytesRead = lngTotalBytesRead + Len(strLineIn) + 2      ' Add 2 bytes to account for CrLf at end of line
                        End If
                    Loop
                Else
                    .SpectraCount = .SpectraCount + 1
                    If .SpectraCount >= lngDtaTxtSpectraDimCount Then
                        lngDtaTxtSpectraDimCount = lngDtaTxtSpectraDimCount + 1000
                        ReDim Preserve .Spectra(lngDtaTxtSpectraDimCount)
                    End If
                    
                    With .Spectra(.SpectraCount - 1)
                        .ParentIon = dblParentIonMass
                        .ParentIonCharge = intChargeFromHeader
                        .ScanNumberStart = lngScanNumberStart
                        .ScanNumberEnd = lngScanNumberEnd
                        .DataCount = 0
                        lngDataDimCount = 0
                        ReDim .XVals(lngDataDimCount)
                        ReDim .YVals(lngDataDimCount)
                    End With
                End If
            End With
        ElseIf IsNumeric(Left(strLineIn, 1)) Then
            ' Add to the most recent scan's data
            
            intParseCount = ParseString(strLineIn, strParsedVals(), 3, " ", "", True, True, False)
            If intParseCount >= 2 Then
                With mDtaTxtData
                    If .SpectraCount > 0 Then
                        With .Spectra(.SpectraCount - 1)
                            .XVals(.DataCount) = Val(strParsedVals(0))
                            .YVals(.DataCount) = Val(strParsedVals(1))
                            .DataCount = .DataCount + 1
                            
                            If .DataCount >= lngDataDimCount Then
                                lngDataDimCount = lngDataDimCount + 100
                                ReDim Preserve .XVals(lngDataDimCount)
                                ReDim Preserve .YVals(lngDataDimCount)
                            End If
                        End With
                    End If
                End With
            End If
        End If

    Loop

    Close #InFileNum

    With mDtaTxtData
        cboDTAScanNumber.Clear
        If .SpectraCount > 0 Then
            For lngIndex = 0 To .SpectraCount - 1
                cboDTAScanNumber.AddItem .Spectra(lngIndex).ScanNumberStart
            Next lngIndex
            
            cboDTAScanNumber.ListIndex = 0
            
            .Initialized = True
            Me.Show
        Else
            .Initialized = False
            Me.Hide
        End If
    End With
    
    frmProgress.HideForm
    Exit Sub
    
ReadDtaTxtFileErrorHandler:
    MsgBox "Error reading input file " & strFilePath & vbCrLf & Err.Description & vbCrLf & "Aborting.", vbExclamation + vbOKOnly, "Error"
    frmProgress.HideForm
    
End Sub

Private Sub ShowDtaTextSpectrum()
    Dim lngSelectedIndex As Long
    
    If cboDTAScanNumber.ListIndex < 0 Then Exit Sub
    
    lngSelectedIndex = cboDTAScanNumber.ListIndex
    
    With mDtaTxtData
        If .Initialized Then
            With .Spectra(lngSelectedIndex)
                frmFragmentationModelling.SetIonMatchList .XVals(), .YVals(), .DataCount, mDtaTxtData.DtaTxtFilePath, .ScanNumberStart, .ScanNumberEnd, .ParentIon, .ParentIonCharge
                            
                txtParentIonMass = .ParentIon
                txtParentIonCharge = .ParentIonCharge
                lblScanNumberEnd = .ScanNumberEnd
            End With
        End If
    End With
        
    On Error Resume Next
    
    cboDTAScanNumber.SetFocus

End Sub

Public Sub ToggleAlwaysOnTop(blnStayOnTop As Boolean)
    
    Static blnUpdating As Boolean
    
    If blnUpdating Then Exit Sub
    
    Me.ScaleMode = vbTwips
    
    WindowStayOnTop Me.hwnd, blnStayOnTop, Me.ScaleX(Me.Left, vbTwips, vbPixels), Me.ScaleY(Me.Top, vbTwips, vbPixels), Me.ScaleX(Me.Width, vbTwips, vbPixels), Me.ScaleY(Me.Height, vbTwips, vbPixels)
    
    blnUpdating = True
    SetCheckBox chkWindowStayOnTop, blnStayOnTop
    blnUpdating = False
    
End Sub

Private Sub cboDTAScanNumber_Click()
    ShowDtaTextSpectrum
End Sub

Private Sub chkWindowStayOnTop_Click()
    ToggleAlwaysOnTop cChkBox(chkWindowStayOnTop)
End Sub

Private Sub cmdJumpToScan_Click()
    JumpToScanNumber
End Sub

Private Sub cmdOK_Click()
    Me.Hide
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift Or vbCtrlMask Then
        If KeyCode = vbKeyJ Then
            JumpToScanNumber
        End If
    End If
End Sub

Private Sub Form_Load()
    ToggleAlwaysOnTop False
    
    SizeAndCenterWindow Me, cWindowBottomCenter, 10600, 1725
    
    lblScanNumberEnd.Left = cboDTAScanNumber.Left + 60
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    QueryUnloadFormHandler Me, Cancel, UnloadMode
End Sub

