VERSION 5.00
Begin VB.Form frmFinderOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Finder Options"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5955
   HelpContextID   =   3050
   Icon            =   "frmFinderOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   5955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Tag             =   "10800"
   Begin VB.ComboBox cboSearchType 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Tag             =   "10840"
      ToolTipText     =   "Choose converting between different amounts or molarity calculation in a solvent"
      Top             =   2640
      Width           =   2295
   End
   Begin VB.ComboBox cboSortResults 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Tag             =   "10850"
      ToolTipText     =   "Resorts the results the list."
      Top             =   2160
      Width           =   2295
   End
   Begin VB.CheckBox chkVerifyHydrogens 
      Caption         =   "&Smart H atoms"
      Height          =   255
      Left            =   2040
      TabIndex        =   8
      Tag             =   "10950"
      Top             =   1680
      Value           =   1  'Checked
      Width           =   2415
   End
   Begin VB.CheckBox chkSort 
      Caption         =   "So&rt Results"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Tag             =   "10940"
      Top             =   1680
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.CheckBox chkAutoSetBounds 
      Caption         =   "&Automatically adjust Min and Max in bounded search."
      Height          =   615
      Left            =   2880
      TabIndex        =   11
      Tag             =   "10960"
      Top             =   2160
      Value           =   1  'Checked
      Width           =   2415
   End
   Begin VB.TextBox txtChargeMax 
      Height          =   285
      Left            =   5280
      TabIndex        =   4
      Tag             =   "10830"
      Text            =   "4"
      ToolTipText     =   "Maximum charge to limit compounds to"
      Top             =   960
      Width           =   375
   End
   Begin VB.TextBox txtChargeMin 
      Height          =   285
      Left            =   4560
      TabIndex        =   3
      Tag             =   "10820"
      Text            =   "-4"
      ToolTipText     =   "Minimum charge to limit compounds to"
      Top             =   960
      Width           =   375
   End
   Begin VB.CheckBox chkLimitChargeRange 
      Caption         =   "Limit Charge &Range"
      Height          =   255
      Left            =   2040
      TabIndex        =   2
      Tag             =   "10910"
      Top             =   960
      Width           =   2415
   End
   Begin VB.CheckBox chkFindTargetMtoZ 
      Caption         =   "Find &Target m/z"
      Height          =   255
      Left            =   2040
      TabIndex        =   6
      Tag             =   "10930"
      Top             =   1320
      Width           =   2415
   End
   Begin VB.CheckBox chkFindCharge 
      Caption         =   "Find &Charge"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Tag             =   "10900"
      Top             =   960
      Width           =   1815
   End
   Begin VB.CheckBox chkFindMtoZ 
      Caption         =   "Find m/&z"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Tag             =   "10920"
      Top             =   1320
      Width           =   1815
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   480
      Left            =   4560
      TabIndex        =   0
      Tag             =   "4010"
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblMax 
      Caption         =   "Max"
      Height          =   255
      Left            =   5280
      TabIndex        =   14
      Tag             =   "10070"
      Top             =   720
      Width           =   735
   End
   Begin VB.Label lblMin 
      Caption         =   "Min"
      Height          =   255
      Left            =   4560
      TabIndex        =   13
      Tag             =   "10060"
      Top             =   720
      Width           =   615
   End
   Begin VB.Label lblInstructions 
      Caption         =   "Use the checkboxes to select various options for the Formula Finder."
      Height          =   615
      Left            =   120
      TabIndex        =   12
      Tag             =   "10810"
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "frmFinderOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mSortModeMapCount As Integer
Private mSortModeMap() As smcFinderResultsSortModeConstants

Public Function LookupResultsSortModeTypeByIndex(intListIndex As Integer) As smcFinderResultsSortModeConstants
    If intListIndex < 0 Or intListIndex > mSortModeMapCount - 1 Then
        LookupResultsSortModeTypeByIndex = smcSortByFormula
    Else
        LookupResultsSortModeTypeByIndex = mSortModeMap(intListIndex)
    End If
End Function

Private Sub PositionFormControls()
    
    ReDim mSortModeMap(5)
    
    PopulateComboBox cboSearchType, True, "Thorough Search|Bounded Search", 0     ' 10840
    
    frmFinderOptions.cboSortResults.ListIndex = 0

    txtChargeMin.Visible = False
    txtChargeMax.Visible = False
    lblMin.Visible = False
    lblMax.Visible = False
    chkFindTargetMtoZ.Enabled = False
    chkVerifyHydrogens.Enabled = False

End Sub

Public Sub UpdateCheckBoxes()
    Dim blnShowControl As Boolean
    Dim intListIndexSave As Integer

    If cChkBox(chkFindCharge) Then
        blnShowControl = True
    Else
        blnShowControl = False
        chkFindMtoZ.value = vbUnchecked
        chkLimitChargeRange.value = vbUnchecked
    End If
    
    ' Show or hide the FindMtoZ and LimitChargeRange buttons depending on FindCharge's value
    If frmFinder.optType(0).value = True Then
        chkFindMtoZ.Enabled = blnShowControl
    Else
        chkFindMtoZ.Enabled = False
        chkFindMtoZ.value = vbUnchecked
    End If
    chkLimitChargeRange.Enabled = blnShowControl

    If blnShowControl = True And cChkBox(chkLimitChargeRange) Then
        blnShowControl = True
    Else
        blnShowControl = False
    End If
    
    ' If FindCharge is checked and LimitChargeRange is checked then show the Min and Max labels and boxes
    txtChargeMin.Visible = blnShowControl
    txtChargeMax.Visible = blnShowControl
    lblMin.Visible = blnShowControl
    lblMax.Visible = blnShowControl
    
    ' If LimitChargeRange is checked and FindMtoZ is checked then show FindTargetMtoZ
    If blnShowControl = True And cChkBox(chkFindMtoZ) Then
        chkFindTargetMtoZ.Enabled = True
    Else
        chkFindTargetMtoZ.Enabled = False
        chkFindTargetMtoZ.value = vbUnchecked
    End If

    ' If SortResults is checked, then show sort mode
    If cChkBox(chkSort) Then
        cboSortResults.Enabled = True
    Else
        cboSortResults.Enabled = False
    End If
    
    ' If cboSearchMode is bounded search, then show the auto adjust min and max checkbox
    If cboSearchType.ListIndex = 1 Then
        chkAutoSetBounds.Enabled = True
    Else
        chkAutoSetBounds.Enabled = False
    End If
    
    ' Set the lblMWT Caption based on the FindMtoZ box
    If cChkBox(chkFindTargetMtoZ) Then
        frmFinder.lblMWT.Caption = LookupLanguageCaption(10870, "Mass/Charge Ratio of Target") & ":"
    Else
        frmFinder.lblMWT.Caption = LookupLanguageCaption(10875, "Molecular Weight of Target") & ":"
    End If
    
    ' Update the values in the sorttype combo box
    With cboSortResults
        intListIndexSave = .ListIndex
        .Clear
    End With

    ' Clear mSortModeMap
    mSortModeMapCount = 0
    ReDim mSortModeMap(6)
    
    ' Add items to sort box
    AddSortResultsItem LookupLanguageCaption(10850, "Sort by Formula"), smcSortByFormula
        
    If cChkBox(chkFindCharge) Then
        AddSortResultsItem LookupLanguageCaption(10855, "Sort by Charge"), smcSortByCharge
    End If
        
    AddSortResultsItem LookupLanguageCaption(10860, "Sort by MWT"), smcSortByMWT
        
    If cChkBox(chkFindMtoZ) Then
        AddSortResultsItem LookupLanguageCaption(10865, "Sort by m/z"), smcSortByMZ
    End If
        
    AddSortResultsItem LookupLanguageCaption(10867, "Sort by Abs(Delta Mass)"), smcSortByDeltaMass
        
    If intListIndexSave < cboSortResults.ListCount Then
        cboSortResults.ListIndex = intListIndexSave
    Else
        cboSortResults.ListIndex = 0
    End If
    
End Sub

Private Sub AddSortResultsItem(strItemName, eSortMode As smcFinderResultsSortModeConstants)
    cboSortResults.AddItem strItemName
    
    If mSortModeMapCount >= mSortModeMapCount Then
        ReDim Preserve mSortModeMap(UBound(mSortModeMap) * 2)
    End If
    
    mSortModeMap(mSortModeMapCount) = eSortMode
    mSortModeMapCount = mSortModeMapCount + 1
    
End Sub

Private Sub cboSearchType_Click()
    Dim blnBoundedSearch As Boolean, intIndex As Integer
    
    UpdateCheckBoxes
    
    If cboSearchType.ListIndex = 0 Then
        ' Thorough search, so hide the bounds boxes
        blnBoundedSearch = False
    Else
        ' Bounded search, so show the bounds boxes
        blnBoundedSearch = True
    End If
    
    With frmFinder
        ' Show/hide bounds boxes as needed
        For intIndex = 0 To 9
            If cChkBox(.chkElements(intIndex)) Then
                .txtMin(intIndex).Visible = blnBoundedSearch
                .txtMax(intIndex).Visible = blnBoundedSearch
            Else
                .txtMin(intIndex).Visible = False
                .txtMax(intIndex).Visible = False
                .txtPercent(intIndex).Visible = False
            End If
        Next intIndex
    
        If blnBoundedSearch Then
            ' Show labels as needed
            For intIndex = 0 To 9
                If cChkBox(.chkElements(intIndex)) Then Exit For
            Next intIndex
            If intIndex < 10 Then
                ' At least one checked, show labels
                .lblMin.Visible = True
                .lblMax.Visible = True
            End If
        Else
            .lblMin.Visible = False
            .lblMax.Visible = False
        End If
    
        .ResizeForm
    End With
End Sub

Private Sub chkFindCharge_Click()
    UpdateCheckBoxes
End Sub

Private Sub chkFindMtoZ_Click()
    UpdateCheckBoxes

End Sub

Private Sub chkFindTargetMtoZ_Click()
    UpdateCheckBoxes
End Sub

Private Sub chkLimitChargeRange_Click()
    chkLimitChargeRange_Validate False

End Sub

Private Sub chkLimitChargeRange_Validate(Cancel As Boolean)
    UpdateCheckBoxes
End Sub

Private Sub chkSort_Click()
    UpdateCheckBoxes
End Sub

Private Sub cmdOK_Click()
    frmFinderOptions.Hide
End Sub

Private Sub Form_Activate()
    ' Put window in center of screen (and upper third vertically)
    SizeAndCenterWindow Me, cWindowUpperThird, 6050, 3600

    UpdateCheckBoxes
End Sub

Private Sub Form_Click()
    UpdateCheckBoxes
End Sub

Private Sub Form_Load()
    PositionFormControls
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    QueryUnloadFormHandler Me, Cancel, UnloadMode
End Sub

Private Sub txtChargeMax_GotFocus()
    HighlightOnFocus txtChargeMax

End Sub

Private Sub txtChargeMax_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtChargeMax, KeyAscii, True, False, True
End Sub

Private Sub txtChargeMax_Validate(Cancel As Boolean)
    ' Make sure the value is reasonable
    ValidateDualTextBoxes txtChargeMin, txtChargeMax, False, -20, 20, 1
End Sub

Private Sub txtChargeMin_GotFocus()
    HighlightOnFocus txtChargeMin
End Sub

Private Sub txtChargeMin_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtChargeMin, KeyAscii, True, False, True
End Sub

Private Sub txtChargeMin_Validate(Cancel As Boolean)
    ' Make sure the value is reasonable
    ValidateDualTextBoxes txtChargeMin, txtChargeMax, True, -20, 20, 1
End Sub
