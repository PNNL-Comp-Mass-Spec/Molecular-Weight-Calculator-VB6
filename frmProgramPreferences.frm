VERSION 5.00
Begin VB.Form frmProgramPreferences 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Molecular Weight Calculator Preferences"
   ClientHeight    =   6705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9975
   HelpContextID   =   4006
   Icon            =   "frmProgramPreferences.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   9975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Tag             =   "11500"
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   480
      HelpContextID   =   4006
      Left            =   6120
      TabIndex        =   38
      Tag             =   "4020"
      Top             =   6000
      Width           =   1095
   End
   Begin VB.Frame fraFormulaEntry 
      BorderStyle     =   0  'None
      Caption         =   "Formula Entry Options"
      Height          =   5775
      Left            =   120
      TabIndex        =   41
      Top             =   120
      Width           =   9735
      Begin VB.ComboBox cboStartupModule 
         Height          =   315
         Left            =   4680
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Tag             =   "11895"
         ToolTipText     =   "Module to show when the program starts"
         Top             =   4920
         Width           =   4095
      End
      Begin VB.ComboBox cboMaximumFormulasToShow 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Tag             =   "11880"
         ToolTipText     =   $"frmProgramPreferences.frx":08CA
         Top             =   5280
         Width           =   735
      End
      Begin VB.CheckBox chkShowCaution 
         Caption         =   "Show Caution Statements (F7)"
         Height          =   255
         Left            =   4680
         TabIndex        =   28
         Tag             =   "11830"
         Top             =   3240
         Value           =   1  'Checked
         Width           =   4935
      End
      Begin VB.CheckBox chkShowQuickSwitch 
         Caption         =   "Show Element Mode &Quick Switch"
         Height          =   255
         Left            =   4680
         TabIndex        =   29
         Tag             =   "11840"
         Top             =   3600
         Value           =   1  'Checked
         Width           =   4935
      End
      Begin VB.CheckBox chkShowToolTips 
         Caption         =   "Show Tool Tips"
         Height          =   255
         Left            =   4680
         TabIndex        =   30
         Tag             =   "11850"
         Top             =   3960
         Value           =   1  'Checked
         Width           =   3975
      End
      Begin VB.CheckBox chkHideInactiveForms 
         Caption         =   "&Hide inactive program windows"
         Height          =   255
         Left            =   4680
         TabIndex        =   34
         Tag             =   "11870"
         Top             =   5325
         Width           =   4935
      End
      Begin VB.CheckBox chkHighlightTextFields 
         Caption         =   "Highlight Text Fields when Selected"
         Height          =   255
         Left            =   4680
         TabIndex        =   31
         Tag             =   "11860"
         Top             =   4320
         Width           =   4935
      End
      Begin VB.CheckBox chkAutosaveValues 
         Caption         =   "&Autosave options, values, and formulas on exit"
         Height          =   255
         Left            =   4680
         TabIndex        =   27
         Tag             =   "11820"
         Top             =   2880
         Width           =   4935
      End
      Begin VB.Frame fraExitProgramOptions 
         Caption         =   "Exit Program Options"
         Height          =   1335
         Left            =   4560
         TabIndex        =   22
         Tag             =   "11550"
         Top             =   1440
         Width           =   5055
         Begin VB.OptionButton optExitConfirmation 
            Caption         =   "Ignore escape key and do not confirm exit"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   26
            Tag             =   "11715"
            Top             =   960
            Width           =   4755
         End
         Begin VB.OptionButton optExitConfirmation 
            Caption         =   "Ignore escape key but confirm exit"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   25
            Tag             =   "11710"
            Top             =   720
            Width           =   4755
         End
         Begin VB.OptionButton optExitConfirmation 
            Caption         =   "Exit on Escape without confirmation"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   24
            Tag             =   "11705"
            Top             =   480
            Width           =   4755
         End
         Begin VB.OptionButton optExitConfirmation 
            Caption         =   "Exit on Escape with confirmation"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   23
            Tag             =   "11700"
            Top             =   240
            Value           =   -1  'True
            Width           =   4755
         End
      End
      Begin VB.Frame fraFormulaFinderWeightModeWarning 
         Caption         =   "Advanced Tools Weight Mode Options"
         Height          =   1335
         Left            =   4560
         TabIndex        =   19
         Tag             =   "11540"
         Top             =   0
         Width           =   5055
         Begin VB.CheckBox chkAlwaysSwitchToIsotopic 
            Caption         =   "Always switch to &Isotopic Mode automatically"
            Height          =   495
            Left            =   240
            TabIndex        =   20
            Tag             =   "11800"
            Top             =   240
            Width           =   4575
         End
         Begin VB.CheckBox chkNeverShowFormulaFinderWarning 
            Caption         =   "&Never show the weight mode warning dialog"
            Height          =   495
            Left            =   240
            TabIndex        =   21
            Tag             =   "11810"
            Top             =   720
            Width           =   4695
         End
      End
      Begin VB.Frame fraStandardDeviation 
         Caption         =   "Standard Deviation Mode (F12)"
         Height          =   1335
         Left            =   120
         TabIndex        =   8
         Tag             =   "11530"
         Top             =   2400
         Width           =   3975
         Begin VB.OptionButton optStdDevType 
            Caption         =   "Decimal"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   11
            Tag             =   "11690"
            Top             =   720
            Width           =   3315
         End
         Begin VB.OptionButton optStdDevType 
            Caption         =   "Scientific"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   10
            Tag             =   "11685"
            Top             =   480
            Width           =   3435
         End
         Begin VB.OptionButton optStdDevType 
            Caption         =   "Short"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   9
            Tag             =   "11680"
            Top             =   240
            Value           =   -1  'True
            Width           =   3555
         End
         Begin VB.OptionButton optStdDevType 
            Caption         =   "Off"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   12
            Tag             =   "11695"
            Top             =   960
            Width           =   3435
         End
      End
      Begin VB.Frame fraCaseRecognition 
         Caption         =   "Case Recognition Mode (F4)"
         Height          =   1095
         Left            =   120
         TabIndex        =   4
         Tag             =   "11520"
         Top             =   1200
         Width           =   3975
         Begin VB.OptionButton optConvertType 
            Caption         =   "Convert Case Up"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   5
            Tag             =   "11665"
            Top             =   240
            Value           =   -1  'True
            Width           =   3555
         End
         Begin VB.OptionButton optConvertType 
            Caption         =   "Exact Case"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   6
            Tag             =   "11670"
            Top             =   480
            Width           =   3555
         End
         Begin VB.OptionButton optConvertType 
            Caption         =   "Smart Case"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   7
            Tag             =   "11675"
            Top             =   720
            Width           =   3555
         End
      End
      Begin VB.Frame fraAbbreviations 
         Caption         =   "Abbreviation Mode (F3)"
         Height          =   1095
         Left            =   120
         TabIndex        =   0
         Tag             =   "11510"
         Top             =   0
         Width           =   3975
         Begin VB.OptionButton optAbbrevType 
            Caption         =   "Off"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   3
            Tag             =   "11660"
            Top             =   720
            Width           =   3555
         End
         Begin VB.OptionButton optAbbrevType 
            Caption         =   "Normal + Amino Acids"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   2
            Tag             =   "11655"
            Top             =   480
            Width           =   3555
         End
         Begin VB.OptionButton optAbbrevType 
            Caption         =   "Normal"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   1
            Tag             =   "11650"
            Top             =   240
            Value           =   -1  'True
            Width           =   3555
         End
      End
      Begin VB.CheckBox chkComputeCharge 
         Caption         =   "Compute Char&ge"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Tag             =   "11780"
         Top             =   4920
         Width           =   4455
      End
      Begin VB.CheckBox chkBracketsAsParentheses 
         Caption         =   "Treat &Brackets as Parentheses"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Tag             =   "11760"
         Top             =   4200
         Width           =   4455
      End
      Begin VB.CheckBox chkAutoCopyCurrentMWT 
         Caption         =   "A&uto Copy Current Molecular Weight (Ctrl+U)"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Tag             =   "11770"
         Top             =   4560
         Width           =   4455
      End
      Begin VB.CheckBox chkAdvanceOnCalculate 
         Caption         =   "Ad&vance on Calculate (F9)"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Tag             =   "11750"
         Top             =   3840
         Width           =   4455
      End
      Begin VB.Label lblStartupModule 
         Caption         =   "Module to show at startup"
         Height          =   255
         Left            =   4680
         TabIndex        =   32
         Tag             =   "11890"
         Top             =   4680
         Width           =   4095
      End
      Begin VB.Label lblMaximumFormulasToShow 
         Caption         =   "Maximum number of formulas to display"
         Height          =   255
         Left            =   1080
         TabIndex        =   18
         Tag             =   "11885"
         Top             =   5325
         Width           =   3615
      End
   End
   Begin VB.CommandButton cmdRestoreOptions 
      Caption         =   "&Restore default options"
      Height          =   615
      Left            =   1800
      TabIndex        =   36
      Tag             =   "11610"
      Top             =   6000
      Width           =   1335
   End
   Begin VB.CommandButton cmdSaveOptions 
      Caption         =   "&Save options as defaults"
      Height          =   615
      Left            =   360
      TabIndex        =   35
      Tag             =   "11600"
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Cl&ose"
      Default         =   -1  'True
      Height          =   480
      Left            =   4800
      TabIndex        =   37
      Tag             =   "4000"
      Top             =   6000
      Width           =   1095
   End
   Begin VB.Label lblHiddenDefaultsLoadedStatus 
      Caption         =   "Hidden Defaults Loaded Status"
      Height          =   255
      Left            =   7800
      TabIndex        =   40
      Top             =   6240
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label lblHiddenFormStatus 
      Caption         =   "Hidden Form Loaded Status"
      Height          =   255
      Left            =   7800
      TabIndex        =   39
      Top             =   6000
      Visible         =   0   'False
      Width           =   2055
   End
End
Attribute VB_Name = "frmProgramPreferences"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const STDDEV_INDEX_OFF = 3

Private lclAbbreviationMode As Integer, lclCaseRecognitionMode As Integer
Private lclStandardDeviationMode As Integer, lclExitProgramOption As Integer
Private lclShowCaution As Integer, lclComputeCharge As Integer
Private lclAdvanceOnCalculate As Integer, lclQuickSwitch As Integer
Private lclAutoSaveValues As Integer, lclAlwaysSwitchToIsotopicMode As Integer
Private lclHideInactiveForms As Integer, lclStartupModule As Integer
Private lclBracketsAsParentheses As Integer, lclAutoCopyCurrentMWT As Integer
Private lclNeverShowFFWarning As Integer, lclShowTooltips As Integer
Private lclHighlightTextFields As Integer, lclMaximumFormulasToShow As Integer

Private Sub PopulateComboBoxes()
    Dim intIndex As Integer, intMaximumFormulaIndexToAllow As Integer
    Const intMinimumFormulaIndexToAllow = 2
    
    intMaximumFormulaIndexToAllow = DetermineMaxAllowableFormulaIndex()
    With cboMaximumFormulasToShow
        .Clear
        For intIndex = 0 To intMaximumFormulaIndexToAllow - intMinimumFormulaIndexToAllow
            .AddItem CStr(intIndex + intMinimumFormulaIndexToAllow + 1)
        Next intIndex
        .ListIndex = .ListCount - 1
    End With
    
    UpdateDynamicComboBox
End Sub

Private Sub RestorePreviousOptions()
    optAbbrevType(lclAbbreviationMode).value = True
    optConvertType(lclCaseRecognitionMode).value = True
    optStdDevType(lclStandardDeviationMode).value = True
    optExitConfirmation(lclExitProgramOption).value = True
    
    chkAlwaysSwitchToIsotopic = lclAlwaysSwitchToIsotopicMode
    chkNeverShowFormulaFinderWarning = lclNeverShowFFWarning
    
    chkAdvanceOnCalculate = lclAdvanceOnCalculate
    chkBracketsAsParentheses = lclBracketsAsParentheses
    chkAutoCopyCurrentMWT = lclAutoCopyCurrentMWT
    chkComputeCharge = lclComputeCharge
    
    chkAutosaveValues = lclAutoSaveValues
    chkShowCaution = lclShowCaution
    chkShowQuickSwitch = lclQuickSwitch
    chkShowToolTips = lclShowTooltips
    chkHighlightTextFields = lclHighlightTextFields
    chkHideInactiveForms = lclHideInactiveForms
    cboStartupModule.ListIndex = lclStartupModule
    cboMaximumFormulasToShow.ListIndex = lclMaximumFormulasToShow

End Sub

Private Sub PositionFormControls()
    Dim intCheckBoxSpacing As Integer
    
    intCheckBoxSpacing = 360
    
    fraFormulaEntry.Top = 120
    fraFormulaEntry.Left = 120
    
    fraAbbreviations.Top = 0
    fraAbbreviations.Left = 120
    
    fraCaseRecognition.Top = 1200
    fraCaseRecognition.Left = fraAbbreviations.Left
    
    fraStandardDeviation.Top = 2400
    fraStandardDeviation.Left = fraAbbreviations.Left
    
    fraFormulaFinderWeightModeWarning.Top = fraAbbreviations.Top
    fraFormulaFinderWeightModeWarning.Left = 4560
    
    fraExitProgramOptions.Top = 1440
    fraExitProgramOptions.Left = fraFormulaFinderWeightModeWarning.Left
    
    chkAdvanceOnCalculate.Top = 3840
    chkAdvanceOnCalculate.Left = 240
    chkBracketsAsParentheses.Top = chkAdvanceOnCalculate.Top + intCheckBoxSpacing * 1
    chkBracketsAsParentheses.Left = chkAdvanceOnCalculate.Left
    chkAutoCopyCurrentMWT.Top = chkAdvanceOnCalculate.Top + intCheckBoxSpacing * 2
    chkAutoCopyCurrentMWT.Left = chkAdvanceOnCalculate.Left
    chkComputeCharge.Top = chkAdvanceOnCalculate.Top + intCheckBoxSpacing * 3
    chkComputeCharge.Left = chkAdvanceOnCalculate.Left
    
    cboMaximumFormulasToShow.Top = 5280
    cboMaximumFormulasToShow.Left = chkAdvanceOnCalculate.Left
    lblMaximumFormulasToShow.Top = 5325
    lblMaximumFormulasToShow.Left = 1080
    
    chkAutosaveValues.Top = 2880
    chkAutosaveValues.Left = 4680
    chkShowCaution.Top = chkAutosaveValues.Top + intCheckBoxSpacing * 1
    chkShowCaution.Left = chkAutosaveValues.Left
    chkShowQuickSwitch.Top = chkAutosaveValues.Top + intCheckBoxSpacing * 2
    chkShowQuickSwitch.Left = chkAutosaveValues.Left
    chkShowToolTips.Top = chkAutosaveValues.Top + intCheckBoxSpacing * 3
    chkShowToolTips.Left = chkAutosaveValues.Left
    chkHighlightTextFields.Top = chkAutosaveValues.Top + intCheckBoxSpacing * 4
    chkHighlightTextFields.Left = chkAutosaveValues.Left
    
    lblStartupModule.Top = 4680
    lblStartupModule.Left = chkAutosaveValues.Left
    cboStartupModule.Top = 4920
    cboStartupModule.Left = chkAutosaveValues.Left
    
    chkHideInactiveForms.Top = 5325
    chkHideInactiveForms.Left = chkAutosaveValues.Left
    
    cmdSaveOptions.Top = 5970
    cmdSaveOptions.Left = 360
    cmdRestoreOptions.Top = cmdSaveOptions.Top
    cmdRestoreOptions.Left = 1800
    
    CmdOK.Top = 6050
    CmdOK.Left = 4800
    cmdCancel.Top = CmdOK.Top
    cmdCancel.Left = 6120
    
End Sub

Private Sub SaveCurrentOptions()
    Dim intIndex As Integer

    For intIndex = 0 To 2
        If optAbbrevType(intIndex).value = True Then
            lclAbbreviationMode = intIndex
        End If
    Next intIndex
 
    For intIndex = 0 To 2
        If optConvertType(intIndex).value = True Then
            lclCaseRecognitionMode = intIndex
        End If
    Next intIndex

    For intIndex = 0 To 3
        If optStdDevType(intIndex).value = True Then
            lclStandardDeviationMode = intIndex
        End If
    Next intIndex

    For intIndex = 0 To 3
        If optExitConfirmation(intIndex).value = True Then
            lclExitProgramOption = intIndex
        End If
    Next intIndex

    lclAlwaysSwitchToIsotopicMode = chkAlwaysSwitchToIsotopic
    lclNeverShowFFWarning = chkNeverShowFormulaFinderWarning
    
    lclAdvanceOnCalculate = chkAdvanceOnCalculate
    lclBracketsAsParentheses = chkBracketsAsParentheses
    lclAutoCopyCurrentMWT = chkAutoCopyCurrentMWT
    lclComputeCharge = chkComputeCharge
    
    lclAutoSaveValues = chkAutosaveValues
    lclShowCaution = chkShowCaution
    lclQuickSwitch = chkShowQuickSwitch
    lclShowTooltips = chkShowToolTips
    lclHighlightTextFields = chkHighlightTextFields
    lclHideInactiveForms = chkHideInactiveForms
    lclStartupModule = cboStartupModule.ListIndex
    lclMaximumFormulasToShow = cboMaximumFormulasToShow.ListIndex
    
End Sub

Public Sub SwapCheck(ThisCheckBoxControl As CheckBox)
        
    With ThisCheckBoxControl
        If .value = vbUnchecked Then
            .value = vbChecked
        Else
            .value = vbUnchecked
        End If
    End With
End Sub

Private Sub cboMaximumFormulasToShow_Validate(Cancel As Boolean)
    Dim intDesiredMaxIndex As Integer
    
    intDesiredMaxIndex = Val(cboMaximumFormulasToShow.List(cboMaximumFormulasToShow.ListIndex)) - 1
    
    If intDesiredMaxIndex = gMaxFormulaIndex Then Exit Sub
    
    ' Only chage intDesiredMaxIndex if the new, desired index is larger than the current value
    ' Reason: it is easy to load new objects, but more difficult to unload them
    If intDesiredMaxIndex > frmMain.GetTopFormulaIndex Then
        gMaxFormulaIndex = intDesiredMaxIndex
    End If
    
    SaveSingleDefaultOption "MaximumFormulasToShow", FormatMaximumFormulasToShowString(intDesiredMaxIndex)

End Sub

Private Sub UpdateDynamicComboBox()
    PopulateComboBox cboStartupModule, True, LookupLanguageCaption(11898, "Main Window|Formula Finder|Capillary Flow Calculator|Mole/Mass Converter|Peptide Sequence Fragmentation Modeller|Amino Acid Notation Converter|Isotopic Distribution Modeller")
    
    If cboStartupModule.ListIndex < 0 Then cboStartupModule.ListIndex = 0
    
End Sub

Private Sub chkAlwaysSwitchToIsotopic_Click()
    If cChkBox(chkAlwaysSwitchToIsotopic) Then
        chkNeverShowFormulaFinderWarning.value = vbChecked
        chkNeverShowFormulaFinderWarning.Enabled = False
    Else
        chkNeverShowFormulaFinderWarning.value = vbUnchecked
        chkNeverShowFormulaFinderWarning.Enabled = True
    End If
    
End Sub

Private Sub chkAutoSaveValues_Click()
    Dim strValueToSave As String
    
    If cChkBox(chkAutosaveValues) Then
        strValueToSave = "1"
    Else
        strValueToSave = "0"
    End If
    
    SaveSingleDefaultOption "AutoSaveValues", strValueToSave
    frmMain.lblStatus.ForeColor = vbWindowText
    frmMain.lblStatus.Caption = LookupLanguageCaption(11900, "Autosave values option saved.")

End Sub

Private Sub chkBracketsAsParentheses_Click()
    objMwtWin.BracketsTreatedAsParentheses = cChkBox(chkBracketsAsParentheses)
    
End Sub

Private Sub chkShowQuickSwitch_Click()
    
    frmMain.ShowHideQuickSwitch chkShowQuickSwitch.value
    
    frmMain.ResizeFormMain True
    
End Sub

Private Sub chkShowToolTips_Click()
    SwitchTips cChkBox(chkShowToolTips)

End Sub

Private Sub cmdCancel_Click()
    RestorePreviousOptions
    frmProgramPreferences.Hide
    lblHiddenFormStatus = "Cancelled"

End Sub

Private Sub cmdOK_Click()
    frmProgramPreferences.Hide
    lblHiddenFormStatus = "Hidden"
End Sub

Private Sub cmdRestoreOptions_Click()
    LoadDefaultOptions True
    
    frmMain.lblStatus.ForeColor = vbWindowText
    frmMain.lblStatus.Caption = LookupLanguageCaption(11910, "Default options restored.")

    If frmProgramPreferences.Visible = True Then
        frmProgramPreferences.cmdSaveOptions.SetFocus
    End If
    
End Sub

Private Sub cmdSaveOptions_Click()
    SaveDefaultOptions
End Sub

Private Sub Form_Activate()
    
    SizeAndCenterWindow Me, cWindowUpperThird, 10000, 7100

    If lblHiddenFormStatus <> "Saved" Then
        SaveCurrentOptions
        lblHiddenFormStatus = "Saved"
    End If
    
    UpdateDynamicComboBox
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim NewValue As Integer
    
    Select Case KeyCode
    Case vbKeyF1
        ' Windows Help command
    Case vbKeyF2
        ' Copy current formula
    Case vbKeyF3
        ' Change abbreviation recognition mode
        If optAbbrevType(0).value = True Then
            optAbbrevType(1).value = True
        ElseIf optAbbrevType(1).value = True Then
            optAbbrevType(2).value = True
        Else
            optAbbrevType(0).value = True
        End If
        frmMain.LabelStatus
    Case vbKeyF4
        ' Change Case Mode
        If optConvertType(0).value = True Then
            optConvertType(1).value = True
        ElseIf optConvertType(1).value = True Then
            optConvertType(2).value = True
        Else
            optConvertType(0).value = True
        End If
        frmMain.LabelStatus
    Case vbKeyF7
        ' Change caution mode
        SwapCheck chkShowCaution
        frmMain.LabelStatus
    Case vbKeyF9
        ' Change advance mode
        SwapCheck chkAdvanceOnCalculate
        frmMain.LabelStatus
    Case vbKeyF12
        ' Change standard deviation mode
        If optStdDevType(0).value = True Then
            optStdDevType(1).value = True
            NewValue = 1
        ElseIf optStdDevType(1).value = True Then
            optStdDevType(2).value = True
            NewValue = 2
        ElseIf optStdDevType(2).value = True Then
            optStdDevType(3).value = True
            NewValue = 3
        Else
            optStdDevType(0).value = True
            NewValue = 0
        End If
        frmMain.LabelStatus
        frmProgramPreferences.optStdDevType(NewValue).SetFocus
    
    Case Else
        If KeyCode = vbKeyU And (Shift And vbCtrlMask) Then    ' And them in case alt or shift was also accidentally pressed
            If chkAutoCopyCurrentMWT = 1 Then
                chkAutoCopyCurrentMWT = 0
            Else
                chkAutoCopyCurrentMWT = 1
            End If
            KeyCode = 0: Shift = 0
        Else
            ' Let the key pass
        End If
    End Select

End Sub

Private Sub Form_Load()
    
    PositionFormControls
    PopulateComboBoxes
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    QueryUnloadFormHandler Me, Cancel, UnloadMode
End Sub

Private Sub optAbbrevType_Click(Index As Integer)
    
    If frmMain.lblHiddenFormStatus <> "Loading" Then
        With frmMain.lblStatus
            .ForeColor = QBColor(COLOR_WARN)
            .Caption = LookupMessage(800)
            gBlnStatusCaution = True
        End With
        objMwtWin.AbbreviationRecognitionMode = Index
    End If

End Sub

Private Sub optConvertType_Click(Index As Integer)
    objMwtWin.CaseConversionMode = Index
End Sub

Private Sub optExitConfirmation_Click(Index As Integer)
    ' Reset dynamic menu captions
    ' Call AppendShortcutKeysToMenuCaptions, which reset all menu captions
    AppendShortcutKeysToMenuCaptions
End Sub

Private Sub optStdDevType_Click(Index As Integer)
    
    objMwtWin.StdDevMode = Index
    
    If Index = STDDEV_INDEX_OFF Then
        gBlnShowStdDevWithMass = False
    Else
        gBlnShowStdDevWithMass = True
    End If
    
    frmMain.Calculate True, True, True, 0, False, False, True
    
    If frmProgramPreferences.Visible = True Then
        frmProgramPreferences.optStdDevType(Index).SetFocus
    End If
    
End Sub

