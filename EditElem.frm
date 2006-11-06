VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmEditElem 
   Caption         =   "Editing Elements"
   ClientHeight    =   5625
   ClientLeft      =   2655
   ClientTop       =   1650
   ClientWidth     =   5955
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   HelpContextID   =   1010
   Icon            =   "EditElem.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5625
   ScaleWidth      =   5955
   Tag             =   "9200"
   Begin VB.Frame fraControls 
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   600
      TabIndex        =   3
      Top             =   4080
      Width           =   4575
      Begin VB.CommandButton CmdOK 
         Caption         =   "&OK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   120
         TabIndex        =   9
         Tag             =   "4010"
         Top             =   960
         Width           =   1035
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   1680
         TabIndex        =   8
         Tag             =   "4020"
         Top             =   960
         Width           =   1035
      End
      Begin VB.CommandButton cmdReset 
         Caption         =   "&Reset to Defaults"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   3120
         TabIndex        =   7
         Tag             =   "9210"
         ToolTipText     =   "Resets elemental weights to their average weights"
         Top             =   960
         Width           =   1395
      End
      Begin VB.CommandButton cmdAverageMass 
         Caption         =   "Use &Average Atomic Weights"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   0
         TabIndex        =   6
         Tag             =   "9230"
         ToolTipText     =   "Sets all elemental weights to their average weights found in nature"
         Top             =   120
         Width           =   1395
      End
      Begin VB.CommandButton cmdIsotopicMass 
         Caption         =   "Use weight of most common &Isotope"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   1560
         TabIndex        =   5
         Tag             =   "9240"
         ToolTipText     =   "Sets all elemental weights to the weight of the element's most common isotope (for high resolution mass spectrometry)"
         Top             =   120
         Width           =   1395
      End
      Begin VB.CommandButton cmdIntegerMass 
         Caption         =   "Use &Nominal integer weight"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   3120
         TabIndex        =   4
         Tag             =   "9245"
         ToolTipText     =   $"EditElem.frx":08CA
         Top             =   120
         Width           =   1395
      End
   End
   Begin VB.ComboBox cboSortBy 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2400
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Tag             =   "9380"
      Top             =   120
      Width           =   2175
   End
   Begin MSFlexGridLib.MSFlexGrid grdElem 
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Tag             =   "9340"
      ToolTipText     =   "Click to change an element's weight or uncertainty"
      Top             =   480
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   6376
      _Version        =   393216
      Rows            =   17
      Cols            =   3
      FixedCols       =   0
      ScrollBars      =   2
   End
   Begin VB.Label lblSortBy 
      Caption         =   "Sort Elements by:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Tag             =   "9390"
      Top             =   150
      Width           =   2055
   End
End
Attribute VB_Name = "frmEditElem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Form wide array
Private PointerArray(MAX_ELEMENT_INDEX) As Integer
Private mValueChanged As Boolean

Private Sub HandleGridClick()
    Dim strMessage As String
    Dim strSymbol As String
    Dim dblMass As Double, dblUncertainty As Double, sngCharge As Single
    Dim lngError As Long
    
    If grdElem.Col <> 0 Then
        ' Can only edit element weights and uncertainties
        ' The Message of the dialog box.
        strMessage = LookupLanguageCaption(9250, "The elemental weight or uncertainty will be changed to the value you type.  Select Reset to reset to the default value or Cancel to ignore any changes.")
    
        ' Display the dialog box and get user's response.
        With frmChangeValue
            .cmdReset.Caption = LookupLanguageCaption(9220, "&Reset")
            .lblHiddenButtonClickStatus = BUTTON_NOT_CLICKED_YET
            .lblInstructions.Caption = strMessage
            .txtValue = Trim(grdElem.Text)
            .rtfValue.Visible = False
            .txtValue.Visible = True
            
            .Show vbModal
            
            If .lblHiddenButtonClickStatus = BUTTON_NOT_CLICKED_YET Then .lblHiddenButtonClickStatus = BUTTON_CANCEL
        End With
        
        If Not frmChangeValue.lblHiddenButtonClickStatus = BUTTON_CANCEL Then
            If frmChangeValue.lblHiddenButtonClickStatus = BUTTON_RESET Then
                ' Reset element stat to default value
                objMwtWin.ResetElement PointerArray(grdElem.Row), grdElem.Col - 1
            Else
                ' Set individual element value
                ' First, grab the current values
                lngError = objMwtWin.GetElement(PointerArray(grdElem.Row), strSymbol, dblMass, dblUncertainty, sngCharge, 0)
                Debug.Assert lngError = 0
                Debug.Assert Left(grdElem.TextMatrix(grdElem.Row, 0), Len(strSymbol)) = strSymbol
                
                ' Update the desired value
                Select Case grdElem.Col
                Case 1: dblMass = CDblSafe(frmChangeValue.txtValue)
                Case 2: dblUncertainty = CDblSafe(frmChangeValue.txtValue)
                Case 3: sngCharge = CSngSafe(frmChangeValue.txtValue)
                End Select
                
                ' Now update the element
                lngError = objMwtWin.SetElement(strSymbol, dblMass, dblUncertainty, sngCharge)
                Debug.Assert lngError = 0
            End If
            UpdateGrid
            mValueChanged = True
        End If
    End If

End Sub

Public Sub PopulateGrid()
    Dim lngIndex As Integer, lngBasicHeight As Integer
    
    With grdElem
        .Left = 120
        .Top = 480
        .Width = frmEditElem.ScaleWidth - 240
        .Height = 3700
        .Rows = objMwtWin.GetElementCount + 1
        
        lngBasicHeight = TextHeight("123456789gT")
        For lngIndex = 0 To objMwtWin.GetElementCount
            .RowHeight(lngIndex) = lngBasicHeight + 60
        Next lngIndex
        
        .Cols = 4
        .ColWidth(0) = (.Width - 350) * 1 / 4
        .ColWidth(1) = (.Width - 350) * 1 / 4
        .ColWidth(2) = (.Width - 350) * 1 / 4
        .ColWidth(3) = (.Width - 350) * 1 / 4
        
        .TextMatrix(0, 0) = LookupLanguageCaption(9350, "Element")
        .TextMatrix(0, 1) = LookupLanguageCaption(9360, "Weight")
        .TextMatrix(0, 2) = LookupLanguageCaption(9370, "Uncertainty")
        .TextMatrix(0, 3) = LookupLanguageCaption(9150, "Charge")
        
        .ColAlignment(3) = vbLeftJustify

        For lngIndex = 1 To objMwtWin.GetElementCount
            PointerArray(lngIndex) = lngIndex
            
            .TextMatrix(lngIndex, 0) = objMwtWin.GetElementSymbol(PointerArray(lngIndex)) & " (" & PointerArray(lngIndex) & "):"
        Next lngIndex
        .Col = 0
        .Row = 1
    End With

End Sub

Public Sub PositionFormControls()
    Dim lngDesiredValue As Long
    
    If Me.WindowState <> vbMaximized Then
        If Me.Width > 6250 Then Me.Width = 6250
    End If
    
    lblSortBy.Top = 100
    cboSortBy.Top = 60
    
    cmdAverageMass.Top = 120
    
    cmdOK.Top = 960
    cmdOK.Left = cmdAverageMass.Left + (cmdAverageMass.Width - cmdOK.Width) / 2
    
    cmdIsotopicMass.Top = cmdAverageMass.Top
    
    cmdCancel.Top = cmdOK.Top
    cmdCancel.Left = cmdIsotopicMass.Left + (cmdIsotopicMass.Width - cmdCancel.Width) / 2
    
    cmdIntegerMass.Top = cmdAverageMass.Top
    
    cmdReset.Top = cmdOK.Top

    With fraControls
        lngDesiredValue = Me.ScaleWidth / 2 - fraControls.Width / 2
        If lngDesiredValue < 0 Then lngDesiredValue = 0
        .Left = lngDesiredValue
        lngDesiredValue = Me.ScaleHeight - fraControls.Height - 120
        If lngDesiredValue < 2000 Then lngDesiredValue = 2000
        .Top = lngDesiredValue
    End With
    
    With grdElem
        .Top = 480
        .Left = 120
        lngDesiredValue = Me.ScaleWidth - .Left - 240
        If lngDesiredValue < 1000 Then lngDesiredValue = 1000
        .Width = lngDesiredValue
        .Height = fraControls.Top - .Top - 120
    End With
End Sub

Private Sub ResetToAverageMassDefaults()
    Dim eResponse As VbMsgBoxResult

    eResponse = YesNoBox(LookupLanguageCaption(9290, "Are you sure you want to reset all the values to the default Elemental values (average weights)?") & "  " & _
                         LookupLanguageCaption(9300, "If executed, this cannot be canceled."), _
                         LookupLanguageCaption(9295, "Reset to Defaults"))

    ' Evaluate the user's response.
    If eResponse = vbYes Then
        ' Reset checkboxes to cause warning box to reappear in formula finder
        frmProgramPreferences.chkAlwaysSwitchToIsotopic.value = vbUnchecked
        frmProgramPreferences.chkNeverShowFormulaFinderWarning.value = vbUnchecked
        mValueChanged = True
        SwitchWeightModeDiskAccess emAverageMass, True
        
        UpdateControls
    End If

End Sub

Private Sub ReSortElementsInGrid()
    Dim lngIndex As Integer, swapVal As Integer, lngRowSave As Integer
    Dim intSortByListIndex As Integer
    Dim strStorage(MAX_ELEMENT_INDEX) As String
    Dim lngErrorID As Long
    Dim strSymbol As String
    
    ' Using the strStorage array since referencing specific rows in the grid is very processor time consuming
    
    ' Resort elements if necessary
        
    lngRowSave = PointerArray(grdElem.Row)
    grdElem.Col = 1
    grdElem.Row = 1
    
    If grdElem.Text <> "" Then
        
        intSortByListIndex = cboSortBy.ListIndex
        Select Case intSortByListIndex
        Case 0  ' element symbol
            grdElem.Col = 0
        Case 1  ' atomic number
                ' basically unsort, done below
        Case 2  ' uncertainty
            grdElem.Col = 2
        Case Else   ' charge (case 3)
            grdElem.Col = 3
        End Select
        
        If intSortByListIndex = 1 Then
            ' Sort by atomic number
            For lngIndex = 1 To objMwtWin.GetElementCount
                PointerArray(lngIndex) = lngIndex
            Next lngIndex
        Else
        
            ' Change mouse pointer to hourglass
            MousePointer = vbHourglass
            
            ' Copy the data to be sorted into strStorage
            For lngIndex = 1 To objMwtWin.GetElementCount
                grdElem.Row = lngIndex
                strStorage(PointerArray(lngIndex)) = grdElem.Text
            Next lngIndex
            
            ' Sort the elements via a shell sort
            Dim MaxRow As Integer, offset As Integer, limit As Integer, switch As Integer
            Dim Row As Integer
    
            ' Set comparison offset to half the number of records (MAX_ELEMENT_INDEX=103 in this case)
            MaxRow = objMwtWin.GetElementCount + 1
            offset = MaxRow \ 2
    
            Do While offset > 0          ' Loop until offset gets to zero.
                limit = MaxRow - offset
                Do
                    switch = 0         ' Assume no switches at this offset.
    
                    ' Compare elements and switch ones out of order:
                    For Row = 1 To limit - 1
                        If intSortByListIndex = 2 Or intSortByListIndex = 3 Then
                            ' Comparing values
                            If CDblSafe(strStorage(PointerArray(Row))) > _
                               CDblSafe(strStorage(PointerArray(Row + offset))) Then
                                ' Swap the pointerarray values
                                swapVal = PointerArray(Row)
                                PointerArray(Row) = PointerArray(Row + offset)
                                PointerArray(Row + offset) = swapVal
                                switch = Row
                            End If
                        Else
                            If strStorage(PointerArray(Row)) > _
                               strStorage(PointerArray(Row + offset)) Then
                                ' Swap the pointerarray values
                                swapVal = PointerArray(Row)
                                PointerArray(Row) = PointerArray(Row + offset)
                                PointerArray(Row + offset) = swapVal
                                switch = Row
                            End If
                        End If
                    Next Row
    
                    ' Sort on next pass only to where last switch was made:
                    limit = switch - offset
                Loop While switch
    
                ' No switches at last offset, try one half as big:
                offset = offset \ 2
            Loop
            ' Change mouse pointer back to default
            MousePointer = vbDefault
        End If
                
        grdElem.Col = 0
        For lngIndex = 1 To objMwtWin.GetElementCount
            grdElem.Row = lngIndex
            
            lngErrorID = objMwtWin.GetElement(PointerArray(lngIndex), strSymbol, 0, 0, 0, 0)
            grdElem.Text = strSymbol & " (" & PointerArray(lngIndex) & "):"
        Next lngIndex
        
        For lngIndex = 1 To objMwtWin.GetElementCount
            If lngRowSave = PointerArray(lngIndex) Then
                grdElem.Row = lngIndex
                grdElem.TopRow = lngIndex
            End If
        Next lngIndex
        UpdateGrid
    End If

End Sub

Public Sub ResetValChangedToFalse()
    mValueChanged = False
End Sub

Private Sub SwitchWeightModeDiskAccess(eNewElementMode As emElementModeConstants, blnRecreateFile As Boolean)
    ' eNewElementMode = 1 is average weights
    ' eNewElementMode = 2 is isotopic weights
    ' eNewElementMode = 3 is integer weights
    
    If eNewElementMode < emAverageMass Or eNewElementMode > emIntegerMass Then
        eNewElementMode = emAverageMass
    End If
    
    If blnRecreateFile Then
        LoadElements CInt(eNewElementMode), False
    Else
        SwitchWeightMode eNewElementMode
    End If
    
    ' Make sure QuickSwitch Element Mode value is correct
    frmMain.ShowHideQuickSwitch frmProgramPreferences.chkShowQuickSwitch.value

End Sub

Private Sub UpdateControls()
    Select Case objMwtWin.GetElementMode()
    Case 1
        cmdAverageMass.Enabled = False
        cmdIsotopicMass.Enabled = True
        cmdIntegerMass.Enabled = True
    Case 2
        cmdAverageMass.Enabled = True
        cmdIsotopicMass.Enabled = False
        cmdIntegerMass.Enabled = True
    Case 3
        cmdAverageMass.Enabled = True
        cmdIsotopicMass.Enabled = True
        cmdIntegerMass.Enabled = False
    Case Else
        SwitchWeightMode emAverageMass
        cmdAverageMass.Enabled = False
        cmdIsotopicMass.Enabled = True
        cmdIntegerMass.Enabled = True
    End Select

    ' Make sure QuickSwitch Element Mode value is correct
    frmMain.ShowHideQuickSwitch frmProgramPreferences.chkShowQuickSwitch.value
    UpdateGrid

End Sub

Private Sub UpdateGrid()
    Dim intIndex As Integer, lngCurrentRow As Integer, lngCurrentColumn As Integer
    Dim lngError As Long
    Dim strSymbol As String
    Dim dblMass As Double, dblUncertainty As Double, sngCharge As Single
    
    lngCurrentRow = grdElem.Row
    lngCurrentColumn = grdElem.Col

    ' Copy the data into the grid
    For intIndex = 1 To objMwtWin.GetElementCount
        lngError = objMwtWin.GetElement(PointerArray(intIndex), strSymbol, dblMass, dblUncertainty, sngCharge, 0)
        Debug.Assert lngError = 0
        
        With grdElem
            .Col = 1
            .Row = intIndex
            .Text = CStr(dblMass)
            .Col = 2
            .Text = CStr(dblUncertainty)
            .Col = 3
            .Text = CStr(sngCharge)
        End With
    Next intIndex

    ' Re-position cursor
    grdElem.Row = lngCurrentRow
    grdElem.Col = lngCurrentColumn

End Sub

Private Sub cboSortBy_Click()
    ReSortElementsInGrid
End Sub

Private Sub cmdAverageMass_Click()
    Dim eResponse As VbMsgBoxResult

    eResponse = YesNoBox(LookupLanguageCaption(9260, "Are you sure you want to reset all the values to their average elemental weights?") & "  " & _
                         LookupLanguageCaption(9300, "If executed, this cannot be canceled."), _
                         LookupLanguageCaption(9265, "Change to Average Weights"))

    ' Evaluate the user's response.
    If eResponse = vbYes Then
        mValueChanged = True
        SwitchWeightModeDiskAccess emAverageMass, True
        
        UpdateControls
    End If

End Sub

Private Sub cmdCancel_Click()
    Dim eResponse As VbMsgBoxResult

    If mValueChanged Then
        eResponse = YesNoBox(LookupLanguageCaption(9310, "Are you sure you want to lose all changes?"), LookupLanguageCaption(9315, "Closing Edit Elements Box"))
        If eResponse = vbYes Then
            LoadElements 0, False
        Else
            Exit Sub
        End If
    End If

    Me.Hide
    mValueChanged = False

End Sub

Private Sub cmdIsotopicMass_Click()
    Dim eResponse As VbMsgBoxResult

    eResponse = YesNoBox(LookupLanguageCaption(9270, "Are you sure you want to reset all the values to their isotopic elemental weights?") & "  " & _
                        LookupLanguageCaption(9300, "If executed, this cannot be canceled."), _
                        LookupLanguageCaption(9275, "Change to Isotopic Weights"))
    
    ' Evaluate the user's response.
    If eResponse = vbYes Then
        mValueChanged = True
        SwitchWeightModeDiskAccess emIsotopicMass, True
        
        UpdateControls
    End If


End Sub

Private Sub cmdIntegerMass_Click()
    Dim eResponse As VbMsgBoxResult

    eResponse = YesNoBox(LookupLanguageCaption(9280, "Are you sure you want to reset all the values to their integer weights?") & "  " & _
                        LookupLanguageCaption(9300, "If executed, this cannot be canceled."), _
                        LookupLanguageCaption(9285, "Change to Integer Weights"))

    ' Evaluate the user's response.
    If eResponse = vbYes Then
        mValueChanged = True
        SwitchWeightModeDiskAccess emIntegerMass, True
        
        UpdateControls
    End If

End Sub

Private Sub cmdOK_Click()
    mValueChanged = True
    
    SaveElements
    
    mValueChanged = False
    Me.Hide
End Sub

Private Sub cmdReset_Click()
    ResetToAverageMassDefaults
End Sub

Private Sub Form_Activate()
    ' Put window in center of screen
    SizeAndCenterWindow Me, cWindowExactCenter, 6250, 6200
    
    UpdateControls
    
End Sub

Private Sub Form_Load()
    PositionFormControls
    
    PopulateComboBox cboSortBy, True, "Element Symbol|Atomic Number|Uncertainty|Charge", 1
    
    PopulateGrid
    UpdateGrid

    mValueChanged = False
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
        cmdCancel_Click
    End If
End Sub

Private Sub Form_Resize()
    PositionFormControls
End Sub

Private Sub grdElem_Click()
    HandleGridClick
End Sub

Private Sub grdelem_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then grdElem_Click
End Sub
