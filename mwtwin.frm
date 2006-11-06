VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "Molecular Weight Calculator"
   ClientHeight    =   5460
   ClientLeft      =   1890
   ClientTop       =   1890
   ClientWidth     =   6270
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
   HelpContextID   =   600
   Icon            =   "mwtwin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5460
   ScaleWidth      =   6270
   Tag             =   "4900"
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5280
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.OptionButton optElementMode 
      Caption         =   "Inte&ger"
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
      HelpContextID   =   4053
      Index           =   2
      Left            =   4680
      TabIndex        =   13
      Tag             =   "5040"
      Top             =   2640
      Width           =   1395
   End
   Begin VB.OptionButton optElementMode 
      Caption         =   "&Isotopic"
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
      HelpContextID   =   4053
      Index           =   1
      Left            =   4680
      TabIndex        =   12
      Tag             =   "5030"
      Top             =   2400
      Width           =   1395
   End
   Begin VB.OptionButton optElementMode 
      Caption         =   "&Average"
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
      HelpContextID   =   4053
      Index           =   0
      Left            =   4680
      TabIndex        =   11
      Tag             =   "5020"
      Top             =   2160
      Value           =   -1  'True
      Width           =   1395
   End
   Begin VB.Frame fraSingle 
      Height          =   3975
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Visible         =   0   'False
      Width           =   4095
      Begin MSFlexGridLib.MSFlexGrid grdPC 
         Height          =   2775
         Left            =   240
         TabIndex        =   8
         Tag             =   "5200"
         ToolTipText     =   "Click to set or reset a target value"
         Top             =   960
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   4895
         _Version        =   393216
         Rows            =   11
         FixedRows       =   0
         FixedCols       =   0
         ScrollBars      =   0
      End
      Begin RichTextLib.RichTextBox rtfFormulaSingle 
         Height          =   495
         Left            =   1080
         TabIndex        =   10
         Tag             =   "5050"
         ToolTipText     =   "5150"
         Top             =   240
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   873
         _Version        =   393217
         MultiLine       =   0   'False
         TextRTF         =   $"mwtwin.frx":08CA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox txtMWTSingle 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Text            =   "MW="
         Top             =   700
         Width           =   3015
      End
      Begin VB.Label lblFormulaSingle 
         Caption         =   "Formula 1:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   900
      End
   End
   Begin VB.Frame fraMulti 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4215
      Begin RichTextLib.RichTextBox rtfFormula 
         Height          =   495
         Index           =   0
         Left            =   960
         TabIndex        =   9
         Tag             =   "5050"
         ToolTipText     =   "Type the molecular formula here."
         Top             =   240
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   873
         _Version        =   393217
         MultiLine       =   0   'False
         TextRTF         =   $"mwtwin.frx":094E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox txtMWT 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Text            =   "MW="
         Top             =   700
         Width           =   3015
      End
      Begin VB.Label lblFormula 
         Caption         =   "Formula 1:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   900
      End
   End
   Begin VB.CommandButton cmdNewFormula 
      Caption         =   "&New Formula"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   1
      TabStop         =   0   'False
      Tag             =   "5110"
      ToolTipText     =   "Adds a new formula to the list"
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "&Calculate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   0
      TabStop         =   0   'False
      Tag             =   "5100"
      ToolTipText     =   "Determines the molecular weight of the current formula"
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ready"
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
      Left            =   -1440
      TabIndex        =   5
      ToolTipText     =   "Double click the status line to expand it"
      Top             =   4800
      Width           =   8295
   End
   Begin VB.Label lblHiddenFormStatus 
      Caption         =   "Unloaded"
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
      Left            =   5040
      TabIndex        =   15
      Top             =   4200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblQuickSwitch 
      Caption         =   "Quick Switch Element Mode"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   14
      Tag             =   "5010"
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label lblValueForX 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   7
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      HelpContextID   =   1000
      Begin VB.Menu mnuEditElements 
         Caption         =   "Edit &Elements Table"
         HelpContextID   =   1010
      End
      Begin VB.Menu mnuEditAbbrev 
         Caption         =   "Edit &Abbreviations"
         HelpContextID   =   1020
      End
      Begin VB.Menu mnuCalculateFile 
         Caption         =   "&Calculate weights from text file"
         HelpContextID   =   1005
      End
      Begin VB.Menu mnuBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print Results"
         HelpContextID   =   1030
      End
      Begin VB.Menu mnuBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         HelpContextID   =   1040
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      HelpContextID   =   2000
      Begin VB.Menu mnuCut 
         Caption         =   "Cu&t"
         HelpContextID   =   2010
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy"
         HelpContextID   =   2010
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste"
         HelpContextID   =   2010
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete"
         HelpContextID   =   2010
      End
      Begin VB.Menu mnuCopyRTF 
         Caption         =   "Copy Current Formula as &RTF"
         HelpContextID   =   2030
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCopyMWT 
         Caption         =   "Copy Current &Molecular Weight"
         HelpContextID   =   2040
         Shortcut        =   ^W
      End
      Begin VB.Menu mnuCopyPC 
         Caption         =   "Copy P&ercent Composition Data"
         HelpContextID   =   2040
      End
      Begin VB.Menu mnuBar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCopyCurrent 
         Caption         =   "Duplicate Current &Formula"
         HelpContextID   =   2050
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuEraseAll 
         Caption         =   "Erase &All Formulas"
         HelpContextID   =   2060
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuEraseCurrent 
         Caption         =   "Erase Current Formula"
         HelpContextID   =   2070
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuBar9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExpandAbbrev 
         Caption         =   "E&xpand Abbreviations"
         HelpContextID   =   2080
      End
      Begin VB.Menu mnuEmpirical 
         Caption         =   "Convert to Empirical F&ormula"
         HelpContextID   =   2090
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      HelpContextID   =   3000
      Begin VB.Menu mnuViewType 
         Caption         =   "&Muti View"
         Checked         =   -1  'True
         HelpContextID   =   3010
         Index           =   0
      End
      Begin VB.Menu mnuViewType 
         Caption         =   "&Single View"
         HelpContextID   =   3020
         Index           =   1
      End
      Begin VB.Menu mnuBar5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPercentSolver 
         Caption         =   "&Percent Solver"
         HelpContextID   =   3030
         Begin VB.Menu mnuPercentType 
            Caption         =   "O&ff"
            Checked         =   -1  'True
            HelpContextID   =   3030
            Index           =   0
         End
         Begin VB.Menu mnuPercentType 
            Caption         =   "&On"
            HelpContextID   =   3030
            Index           =   1
         End
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      HelpContextID   =   3035
      Begin VB.Menu mnuMMConvert 
         Caption         =   "&Mole/Mass Converter"
         HelpContextID   =   3040
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuFinder 
         Caption         =   "&Formula Finder"
         HelpContextID   =   3050
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuAminoAcidNotationConverter 
         Caption         =   "&Amino Acid Notation Converter"
         HelpContextID   =   3055
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuBar13 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPeptideSequenceFragmentation 
         Caption         =   "&Peptide Sequence Fragmentation Modelling"
         HelpContextID   =   3080
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuIsotopicDistribution 
         Caption         =   "&Isotopic Distribution Modelling"
         HelpContextID   =   3100
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuDisplayIsotopicDistribution 
         Caption         =   "Show Isotopic &Distribution for Current Formula"
         HelpContextID   =   3100
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuBar11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCalculator 
         Caption         =   "Math &Calculator"
         HelpContextID   =   3060
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuCapillaryFlow 
         Caption         =   "Capillar&y Flow Calculator"
         HelpContextID   =   3070
         Shortcut        =   ^Y
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      HelpContextID   =   4000
      Begin VB.Menu mnuChooseLanguage 
         Caption         =   "Choose &Language"
         HelpContextID   =   4003
      End
      Begin VB.Menu mnuProgramOptions 
         Caption         =   "Change Program &Preferences"
         HelpContextID   =   4006
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuChangeFont 
         Caption         =   "Change &Formula Font"
         HelpContextID   =   4060
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuStayOnTop 
         Caption         =   "Stay on &Top"
         HelpContextID   =   4065
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuBar10 
         Caption         =   "-"
         HelpContextID   =   4080
      End
      Begin VB.Menu mnuDefaultsOptions 
         Caption         =   "&Save and Restore Default Values"
         HelpContextID   =   4080
         Begin VB.Menu mnuRestoreValues 
            Caption         =   "&Restore Default Values and Formulas"
            HelpContextID   =   4080
         End
         Begin VB.Menu mnuSaveValues 
            Caption         =   "Save &Values and Formulas Now"
            HelpContextID   =   4080
         End
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      HelpContextID   =   5000
      Begin VB.Menu mnuOverview 
         Caption         =   "&Program Overview"
         HelpContextID   =   5000
      End
      Begin VB.Menu mnuShowTips 
         Caption         =   "&Show Tool Tips"
         Checked         =   -1  'True
         HelpContextID   =   5000
      End
      Begin VB.Menu mnuBar7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About MWT"
         HelpContextID   =   6000
      End
   End
   Begin VB.Menu mnuRightClick 
      Caption         =   "RightClickMenu"
      Begin VB.Menu mnuRightClickUndo 
         Caption         =   "&Undo"
      End
      Begin VB.Menu mnuRightClickSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRightClickCut 
         Caption         =   "Cu&t"
      End
      Begin VB.Menu mnuRightClickCopy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnuRightClickPaste 
         Caption         =   "&Paste"
      End
      Begin VB.Menu mnuRightClickDelete 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnuRightClickSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRightClickSelectAll 
         Caption         =   "Select &All"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const PCT_SOLVER_INITIAL_DIVISION_COUNT = 10
Private Const FORMULA_OFFSET_AMOUNT = 750          ' Distance to adjust screen when adding new formulas

Private Const FORMULA_HIGHLIGHTED = "Highlighted"
Private Const FORM_STAY_ON_TOP_NEWLY_ENABLED = "StayOnTopEnabled"
Private Const MAIN_HEIGHT = 3950                ' Minimum height for the form
Private Const MAIN_HEIGHT_SINGLE = 5400

Private Type udtPercentSolverParametersType
    SolverOn As Boolean
    SolverWorking As Boolean
    XVal As Double
    SolverCounter As Long
    blnMaxReached As Boolean
    DivisionCount As Integer
    IterationStopValue As Integer
    StdDevModeEnabledSaved As Boolean
    
    BestXVal As Double                  ' The x value that gave the lowest residual sum
    BestXSumResiduals As Double         ' The sum of the residuals when using BestXVal
    
End Type
    
Private Type udtMultiLineErrorDescriptionType
    ErrorID As Long                  ' Contains the error number (used in the LookupMessage function).  In addition, if a program error occurs, ErrorParams.ErrorID = -10
    ErrorLine As Integer
    ErrorPosition As Integer
    ErrorCharacter As String
End Type
    
Private Type udtPCompGridValuesType
    Text As String
    value As Double
    Goal As Double
    Locked As Boolean
End Type

Private Type udtPctCompValuesType
    GridCell(MAX_ELEMENT_INDEX) As udtPCompGridValuesType
    TotalElements As Integer        ' Total number of elements whose percent composition values are shown
    LockedValuesCount As Integer    ' The total number of locked percent composition values
End Type


' General Form-Wide Variables
Private mBlnBatchProcess As Boolean
Private mBatchProcessFilename As String
Private mBatchProcessOutfile As String
Private mBatchProcessOverwrite As Boolean
Private mPctCompValues(MAX_FORMULAS) As udtPctCompValuesType            ' 0-based array

Private mViewMode As vmdViewModeConstants
Private mCurrentFormulaIndex As Integer             ' The current formula line (0 through mTopFormulaIndex)
Private mTopFormulaIndex As Integer                 ' Ranges from 0 to mTopFormulaIndex;  Number of formulas is mTopFormulaIndex + 1
Private mStdDevModSaved As Integer
Private mKeyPressAbortPctSolver As Integer
Private mCurrentFormWidth As Integer, mCurrentFormHeight As Integer

Private mUsingPCGrid As Boolean
Private mIgnoreNextGotFocus As Boolean
Private mDifferencesFormShown As Boolean
Private mPercentCompositionGridRowCount As Integer

Private mErrorParams As udtMultiLineErrorDescriptionType

Private PctSolverParams As udtPercentSolverParametersType

Public Function AddNewFormulaWrapper(Optional blnForceAddition As Boolean = False) As Boolean
    ' Returns True if a new formula was added or a blank line was found and selected; false otherwise
    
    Dim intIndex As Integer
    Dim blnSuccess As Boolean
    
On Error GoTo AddNewFormulaWrapperErrorHandler
    
    ' Only call this function if the form is visible, otherwise, things can go wrong
    Debug.Assert Me.Visible
    
    If rtfFormulaSingle.Text = "" And Not blnForceAddition Then
        ' Current formula is already blank
        RaiseLabelStatusError 570
        blnSuccess = False
    Else
        If mViewMode = vmdMultiView Then
            ' Multi view
            If Not blnForceAddition Then
                ' Search for a blank line; move to it if found
                For intIndex = 0 To mTopFormulaIndex
                    If rtfFormula(intIndex).Text = "" Then
                        ' The .SetFocus event can cause an error if a form is shown modally
                        ' Must use Resume Next error handling to avoid an error
                        On Error Resume Next
                        rtfFormula(intIndex).SetFocus
                        blnSuccess = True
                        Exit Function
                    End If
                Next intIndex
            End If
            
            ' Actually add the new formula
            If AddNewFormulaWork(True) Then
                ' The .SetFocus event can cause an error if a form is shown modally
                ' Must use Resume Next error handling to avoid an error
                On Error Resume Next
                rtfFormula(mCurrentFormulaIndex).SetFocus
                blnSuccess = True
            Else
                blnSuccess = False
            End If
        Else
            ' Single view
            If PctSolverParams.SolverOn Then
                ' Percent solver is on; must turn off before adding new formula
                RaiseLabelStatusError 580
                blnSuccess = False
            Else
                If Not blnForceAddition Then
                    ' Look for a blank formula; if found, display it
                    ' Otherwise, add a new one
                    For intIndex = 0 To mTopFormulaIndex
                        If rtfFormula(intIndex).Text = "" Then
                            mCurrentFormulaIndex = intIndex
                            rtfFormulaSingle.Text = rtfFormula(mCurrentFormulaIndex).Text
                            lblFormulaSingle.Caption = ConstructFormulaLabel(mCurrentFormulaIndex)
                            txtMWTSingle.Text = txtMWT(mCurrentFormulaIndex).Text
                            If txtMWTSingle.Text <> "" Then
                                txtMWTSingle.Visible = True
                            Else
                                txtMWTSingle.Visible = False
                            End If
                            LabelStatus
                            Calculate True
                            rtfFormulaSingle.SetFocus
                            blnSuccess = True
                            Exit Function
                        End If
                    Next intIndex
                End If
                
                ' Add a new formula
                If AddNewFormulaWork(False) Then
                    rtfFormulaSingle.SetFocus
                    blnSuccess = True
                Else
                    blnSuccess = False
                End If
            End If
        End If
    End If
    
    AddNewFormulaWrapper = blnSuccess
    Exit Function

AddNewFormulaWrapperErrorHandler:
    Debug.Assert False
    GeneralErrorHandler "frmMain|AddNewFormulaWrapper", Err.Number, Err.Description

End Function

Private Function AddNewFormulaWork(boolMultiView As Boolean) As Boolean
    ' Returns true if new formula actually added
    
On Error GoTo AddNewFormulaWorkErrorHandler

    ' Check if mTopFormulaIndex is < gMaxFormulaIndex
    If mTopFormulaIndex >= gMaxFormulaIndex Then
        ' Already have max formula fields, can't add more
        RaiseLabelStatusError 560
        AddNewFormulaWork = False
        Exit Function
    End If
    
    mTopFormulaIndex = mTopFormulaIndex + 1               ' Increment field count.
    mCurrentFormulaIndex = mTopFormulaIndex

    ' Create new Formula Text Field
    Load rtfFormula(mTopFormulaIndex)
    
    ' Set new Text Field under previous field.
    With rtfFormula(mTopFormulaIndex)
        .Top = rtfFormula(mTopFormulaIndex - 1).Top + FORMULA_OFFSET_AMOUNT
        .TabIndex = mTopFormulaIndex
        .Text = ""
    End With
    cmdCalculate.TabIndex = mTopFormulaIndex + 1
    cmdNewFormula.TabIndex = mTopFormulaIndex + 2
    optElementMode(0).TabIndex = mTopFormulaIndex + 3
    optElementMode(1).TabIndex = mTopFormulaIndex + 4
    optElementMode(2).TabIndex = mTopFormulaIndex + 5
    
    If cChkBox(frmProgramPreferences.chkShowToolTips) Then
        rtfFormula(mTopFormulaIndex).ToolTipText = LookupToolTipLanguageCaption(5050, "Type the molecular formula here")
    End If
    
    ' Create new Formula Label Field
    Load lblFormula(mTopFormulaIndex)
    
    ' Set new Text Field under previous field.
    lblFormula(mTopFormulaIndex).Top = lblFormula(mTopFormulaIndex - 1).Top + FORMULA_OFFSET_AMOUNT
    lblFormula(mTopFormulaIndex).Caption = ConstructFormulaLabel(mTopFormulaIndex)
    lblFormulaSingle.Caption = lblFormula(mTopFormulaIndex).Caption

    ' Create new Molecular Weight Text Field
    Load txtMWT(mTopFormulaIndex)
    
    ' Set new Field under previous field.
    txtMWT(mTopFormulaIndex).Top = txtMWT(mTopFormulaIndex - 1).Top + FORMULA_OFFSET_AMOUNT
    txtMWT(mTopFormulaIndex).Text = ""
    If txtMWTSingle.Text <> "" Then txtMWTSingle.Visible = True Else txtMWTSingle.Visible = False
    
    ' Make appropriate new fields visible
    rtfFormula(mTopFormulaIndex).Visible = True
    lblFormula(mTopFormulaIndex).Visible = True
    txtMWT(mTopFormulaIndex).Visible = False
    
    ' Change window heights and set focus
    fraMulti.Height = fraMulti.Height + FORMULA_OFFSET_AMOUNT
    If boolMultiView Then
        With frmMain
            If .WindowState = vbNormal Then
                ResizeFormMain True
                If .Height + .Top > Screen.Height Then
                    .Top = 0
                End If
            End If
        End With
    Else
        rtfFormulaSingle.Text = rtfFormula(mCurrentFormulaIndex).Text
        lblFormulaSingle.Caption = ConstructFormulaLabel(mCurrentFormulaIndex)
        txtMWTSingle.Text = txtMWT(mCurrentFormulaIndex).Text
        If txtMWTSingle.Text <> "" Then
            txtMWTSingle.Visible = True
        Else
            txtMWTSingle.Visible = False
        End If
    End If

    ' Make sure txtMWT is not a tab stop
    txtMWT(mTopFormulaIndex).TabStop = False

    AddNewFormulaWork = True
    Exit Function
    
AddNewFormulaWorkErrorHandler:
    Debug.Assert False
    GeneralErrorHandler "frmMain|AddNewFormulaWork", Err.Number, Err.Description
    
End Function

Private Sub AddToDiffForm(strAdd As String, intLabelIndex As Integer)
   
    Select Case intLabelIndex
    Case 1
        frmDiff.lblDiff1.Caption = frmDiff.lblDiff1.Caption & vbCrLf & strAdd
    Case 2
        frmDiff.lblDiff2.Caption = frmDiff.lblDiff2.Caption & vbCrLf & strAdd
    Case Else
        frmDiff.lblDiff3.Caption = frmDiff.lblDiff3.Caption & vbCrLf & strAdd
    End Select

End Sub

Public Function Calculate(blnProcessAllFormulas As Boolean, Optional blnUpdateControls As Boolean = True, Optional blnStayOnCurrentFormula As Boolean = False, Optional intSingleFormulaToProcess As Integer = 0, Optional blnConvertToEmpirical As Boolean = False, Optional blnExpandAbbreviations As Boolean = False, Optional blnForceRecalculation As Boolean = False, Optional dblValueForX As Double = 1#, Optional blnSetFocusToFormula As Boolean = True) As Long
    ' Re-calculate the molecular weight and percent composition values
    '  for some or all of the formulas
    ' If blnConvertToEmpirical = True then convert to empirical formulas
    ' If blnExpandAbbreviations = True then expand abbreviations to their elemental equivalents
    
    ' Returns 0 if success; the error ID if an error
    
    Dim intFormulaIndex As Integer, intCurrentFormulaIndexSave As Integer
    Dim intIndexStart As Integer, intIndexEnd As Integer
    Dim intElementIndex As Integer
    Dim intGridInUseCount As Integer
    Dim lngError As Long
    Dim strNewFormula As String, strMWTText As String
    Dim strCautionStatement As String
    Dim strStatus As String
    
    If blnProcessAllFormulas Then
        intIndexStart = 0
        intIndexEnd = mTopFormulaIndex
    Else
        intIndexStart = intSingleFormulaToProcess
        intIndexEnd = intSingleFormulaToProcess
    End If
    
On Error GoTo CalculateErrorHandler

    gBlnErrorPresent = False
    mErrorParams.ErrorID = 0
    mErrorParams.ErrorLine = -1
    mErrorParams.ErrorPosition = -1
    mErrorParams.ErrorCharacter = ""
    
    intCurrentFormulaIndexSave = mCurrentFormulaIndex
    For intFormulaIndex = intIndexStart To intIndexEnd
        
        If Len(Trim(rtfFormula(intFormulaIndex).Text)) = 0 Then
            ' Formula is blank
            objCompounds(intFormulaIndex).SetFormula ""
            
            ' Set txtMWT().Text to "" and hide
            txtMWT(intFormulaIndex).Text = ""
            txtMWT(intFormulaIndex).Visible = False
        Else
            If rtfFormula(intFormulaIndex).Tag = FORMULA_CHANGED Or _
               blnForceRecalculation Or _
               blnConvertToEmpirical Or _
               blnExpandAbbreviations Then
    
                With objCompounds(intFormulaIndex)
                    .ValueForX = dblValueForX
                    ' Note: Setting the formula automatically computes the mass
                    lngError = .SetFormula(rtfFormula(intFormulaIndex).Text)
                    
                    strCautionStatement = .CautionDescription
                    If Len(strCautionStatement) > 0 And cChkBox(frmProgramPreferences.chkShowCaution) Then
                        With lblStatus
                            .Caption = LookupLanguageCaption(22000, "Caution") & ", " & strCautionStatement
                            gBlnStatusCaution = True
                            .ForeColor = QBColor(COLOR_WARN)
                        End With
                    End If
                End With
    
                If blnConvertToEmpirical Then
                    ' Note that Converting to the empirical formula automatically expands abbreviations to their empirical formulas
                    strNewFormula = objCompounds(intFormulaIndex).ConvertToEmpirical
                    Debug.Assert objCompounds(intFormulaIndex).FormulaCapitalized = strNewFormula
                ElseIf blnExpandAbbreviations Then
                    strNewFormula = objCompounds(intFormulaIndex).ExpandAbbreviations
                    Debug.Assert objCompounds(intFormulaIndex).FormulaCapitalized = strNewFormula
                End If
    
                ' Update formula to formatted version
                rtfFormula(intFormulaIndex).TextRTF = objMwtWin.TextToRTF(objCompounds(intFormulaIndex).FormulaCapitalized)
                
                Debug.Assert objMwtWin.ErrorID = lngError
                
                If objMwtWin.ErrorID = 0 Then
                    ' No error
                    ' Reset .Tag
                    rtfFormula(intFormulaIndex).Tag = ""
                    rtfFormulaSingle.Tag = ""
                    
                    If blnUpdateControls Then
                        ' Refresh screen
                        If Not gBlnShowStdDevWithMass Then
                            ' Standard Deviation Mode is Off
                            strMWTText = CStr(objCompounds(intFormulaIndex).Mass(False))
                        Else
                            strMWTText = objCompounds(intFormulaIndex).MassAndStdDevString(False)
                        End If
                        
                        strMWTText = ConstructMWEqualsCaption & strMWTText
                        
                        If cChkBox(frmProgramPreferences.chkComputeCharge) Then
                            strMWTText = strMWTText & "   " & _
                                            LookupLanguageCaption(9150, "Charge") & " = " & _
                                            ChargeValueToString(objCompounds(intFormulaIndex).Charge)
                        End If
                        
                        txtMWT(intFormulaIndex).Text = strMWTText
                        txtMWT(intFormulaIndex).Visible = True
                        
                        If intFormulaIndex = intCurrentFormulaIndexSave Then
                            txtMWTSingle.Text = strMWTText
                            txtMWTSingle.Visible = True
                        End If
                        
                    End If
                
                Else
                    ' Error occurred
                    If lngError <> 0 And mErrorParams.ErrorID = 0 Then
                        With mErrorParams
                            .ErrorID = objMwtWin.ErrorID
                            .ErrorLine = intFormulaIndex
                            .ErrorPosition = objMwtWin.ErrorPosition
                            .ErrorCharacter = objMwtWin.ErrorCharacter
                        End With
                    End If
                    
                    ' Hide txtMWT
                    If mViewMode = vmdMultiView Then
                        txtMWT(intFormulaIndex).Visible = False
                    Else
                        txtMWTSingle.Visible = False
                    End If
                End If
    
            End If
        End If
        
    Next intFormulaIndex
    
    If mViewMode = vmdSingleView Then
        ' Need to display percent composition values
        ' Further, if PctSolver is running, then need to compute additional values
        
        With mPctCompValues(mCurrentFormulaIndex)
            .TotalElements = objCompounds(mCurrentFormulaIndex).GetUsedElementCount
            For intElementIndex = 1 To objMwtWin.GetElementCount
                If objCompounds(mCurrentFormulaIndex).ElementPresent(intElementIndex) Then
                    intGridInUseCount = intGridInUseCount + 1
                    .GridCell(intGridInUseCount).Text = objCompounds(mCurrentFormulaIndex).GetPercentCompositionForElementAsString(intElementIndex, gBlnShowStdDevWithMass)
                    .GridCell(intGridInUseCount).value = objCompounds(mCurrentFormulaIndex).GetPercentCompositionForElement(intElementIndex)
                End If
            Next intElementIndex
        End With

        If objCompounds(mCurrentFormulaIndex).Mass(False) > 0 Then
            ' Only proceed with percent solver computations if a non-zero mass
            
            If mErrorParams.ErrorID = 0 And PctSolverParams.SolverOn Then
                mErrorParams.ErrorID = PerformPercentSolverCalculations(mCurrentFormulaIndex)
                If mErrorParams.ErrorID <> 0 Then
                    With mErrorParams
                        .ErrorLine = mCurrentFormulaIndex
                        .ErrorPosition = objMwtWin.ErrorPosition
                        .ErrorCharacter = objMwtWin.ErrorCharacter
                    End With
                End If
            End If
        Else
            If rtfFormula(mCurrentFormulaIndex).Text = "" Then
                If Not PctSolverParams.SolverOn Then
                    lblValueForX.Caption = ""
                    lblValueForX.Visible = False
                Else
                    lblValueForX.ForeColor = vbWindowText
                    lblValueForX.Caption = LookupLanguageCaption(3750, "% Solver On")
                    lblValueForX.Visible = True
                End If
            End If
        End If

        If blnUpdateControls Then
            If mErrorParams.ErrorID = 0 Then
                UpdateGrid grdPC, mCurrentFormulaIndex
            Else
                grdPC.Clear
            End If
        End If
        
    End If
    
    ' May need to copy the formula from rtfFormula() to rtfFormulaSingle
    rtfFormulaSingle.Text = rtfFormula(mCurrentFormulaIndex).Text
    txtMWTSingle.Text = txtMWT(mCurrentFormulaIndex).Text
    
    If mErrorParams.ErrorID <> 0 Then             ' Inform user of error and location
        lblStatus.ForeColor = QBColor(COLOR_ERR)
        strStatus = LookupMessage(mErrorParams.ErrorID)
        If Len(mErrorParams.ErrorCharacter) > 0 Then
            strStatus = strStatus & ": " & mErrorParams.ErrorCharacter
        End If
        lblStatus.Caption = strStatus
        gBlnErrorPresent = True
        If mErrorParams.ErrorLine >= 0 And mErrorParams.ErrorPosition >= 0 Then
            RTFHighlightCharacter mErrorParams.ErrorLine, mErrorParams.ErrorPosition
        End If
        
        ResizeFormMain False, False
    Else
        ' No error
        
        ' Also, copy current molecular weight value to clipboard if Auto Copy is turned on
        If cChkBox(frmProgramPreferences.chkAutoCopyCurrentMWT) Then
            CopyCurrentMWT True
        End If
    End If
    
    ' See if an x is present after a bracket in the current formula
    ' If yes, update lblValueForX to indicate the x value
    ' Otherwise, hide lblValueForX
    If objCompounds(mCurrentFormulaIndex).XIsPresentAfterBracket Then
        lblValueForX.ForeColor = QBColor(COLOR_SOLVER)
        lblValueForX.Caption = LookupLanguageCaption(3710, "x is") & " " & Format(dblValueForX, "#0.0######")
        lblValueForX.Visible = True
    Else
        lblValueForX.Visible = False
    End If
    
    ' Position Cursor
    If rtfFormula(mCurrentFormulaIndex).Tag = "" And _
       rtfFormula(mCurrentFormulaIndex).Text <> "" And _
       blnUpdateControls Then
        If mViewMode = vmdMultiView Then
            If blnProcessAllFormulas And Not blnStayOnCurrentFormula And _
               cChkBox(frmProgramPreferences.chkAdvanceOnCalculate) Then
                If mCurrentFormulaIndex = mTopFormulaIndex And mTopFormulaIndex <> gMaxFormulaIndex Then
                    AddNewFormulaWrapper
                Else
                    mCurrentFormulaIndex = mCurrentFormulaIndex + 1
                    If mCurrentFormulaIndex > mTopFormulaIndex Then
                        ' This should never happen
                        Debug.Assert False
                        mCurrentFormulaIndex = mTopFormulaIndex
                    End If
                End If
            End If
        End If
        
        If frmMain.Visible And blnUpdateControls And blnSetFocusToFormula Then
            SetFocusToFormulaByIndex mCurrentFormulaIndex
        End If
    End If
    
    Calculate = mErrorParams.ErrorID
    
    Exit Function

CalculateErrorHandler:
    GeneralErrorHandler "frmMain|CalculateMultiLine", Err.Number, Err.Description
    
End Function

Private Sub ClearGrid(ThisGrid As MSFlexGrid)
    Dim elementi As Integer
    
    With mPctCompValues(mCurrentFormulaIndex)
        For elementi = 1 To objMwtWin.GetElementCount
            .GridCell(elementi).Text = ""
            .GridCell(elementi).value = 0
        Next elementi
    End With
    
    UpdateGrid ThisGrid, mCurrentFormulaIndex

End Sub

Private Sub ConvertToEmpirical()
    ' Expand abbreviations of current formula into their elemental representations
    ' Then, rearrange elements to be in the order carbon, then hydrogen, then
    '  the rest alphabetically
    
    Dim eResponse As VbMsgBoxResult
    
    ' Display the dialog box and get user's response.
    If (frmProgramPreferences.optExitConfirmation(exmEscapeKeyConfirmExit).value = True Or frmProgramPreferences.optExitConfirmation(exmIgnoreEscapeKeyConfirmExit).value = True) Then
        eResponse = YesNoBox(LookupLanguageCaption(3610, "Are you sure you want to convert the current formula into its empirical formula?"), _
                            LookupLanguageCaption(3615, "Convert to Empirical Formula"))
    Else
        eResponse = vbYes
    End If

    ' Evaluate the user's response.
    If eResponse = vbYes Then
        Calculate False, True, True, mCurrentFormulaIndex, True
    End If

End Sub

Private Sub DisplayIsotopicDistribution(intFormulaIndexToDisplay As Integer)
    If intFormulaIndexToDisplay <= frmMain.GetTopFormulaIndex Then
        frmIsotopicDistribution.rtfFormula.Text = frmMain.rtfFormula(intFormulaIndexToDisplay).Text
        frmIsotopicDistribution.Show
        frmIsotopicDistribution.StartIsotopicDistributionCalcs True
    End If
End Sub

Private Sub DuplicateCurrentFormula()
    ' Copy current formula to new line or existing blank line
    
    Dim blnFound As Boolean, intIndex As Integer, strCopyText As String
    
    ' Only do this if Percent Solver Mode is not on
    If Not PctSolverParams.SolverOn Then
            If rtfFormula(mCurrentFormulaIndex).Text <> "" Then
            ' Proceed only if the line is not blank
            strCopyText = rtfFormula(mCurrentFormulaIndex).Text
            RemoveHeightAdjustChar strCopyText
            
            ' Search for an empty formula line
            For intIndex = 0 To mTopFormulaIndex
                If rtfFormula(intIndex).Text = "" Then
                    blnFound = True
                    Exit For
                End If
            Next intIndex
            
            If Not blnFound Then
                ' No empty lines
                If mTopFormulaIndex < gMaxFormulaIndex Then
                    ' Add new formula and copy to end
                    If AddNewFormulaWrapper() = True Then
                        mCurrentFormulaIndex = mTopFormulaIndex
                        rtfFormula(mTopFormulaIndex).Text = strCopyText
                    End If
                Else
                    lblStatus.ForeColor = QBColor(COLOR_ERR)
                    lblStatus.Caption = LookupMessage(630) & "  " & LookupMessage(560)
                End If
            Else
                ' Copy to blank formula
                mCurrentFormulaIndex = intIndex
                rtfFormula(intIndex).Text = strCopyText
            End If
            
            If mViewMode = vmdSingleView Then
                rtfFormulaSingle.SetFocus
                rtfFormulaSingle.Text = rtfFormula(mCurrentFormulaIndex).Text
                lblFormulaSingle.Caption = ConstructFormulaLabel(mCurrentFormulaIndex)
                txtMWTSingle.Text = txtMWT(mCurrentFormulaIndex).Text
                If txtMWTSingle.Text <> "" Then
                    txtMWTSingle.Visible = True
                Else
                    txtMWTSingle.Visible = False
                End If
            End If
            
            ' Make sure the .Tag = FORMULA_CHANGED to assure the calculation occurs
            rtfFormula(mCurrentFormulaIndex).Tag = FORMULA_CHANGED
            Calculate True, True, True
            
            LabelStatus
        Else
            lblStatus.ForeColor = QBColor(COLOR_ERR)
            lblStatus.Caption = LookupMessage(630) & "  " & LookupMessage(640)
        End If
    Else
        lblStatus.ForeColor = QBColor(COLOR_ERR)
        lblStatus.Caption = LookupMessage(630) & "  " & LookupMessage(655)
    End If

End Sub

Private Sub CopyPC()
    Dim strCopyText As String, intIndex As Integer
    Dim strPctCompositions(MAX_ELEMENT_INDEX) As String
    Clipboard.Clear

    strCopyText = ""
    With objCompounds(mCurrentFormulaIndex)
        .GetPercentCompositionForAllElements strPctCompositions()
        
        For intIndex = 1 To objMwtWin.GetElementCount
            If Len(strPctCompositions(intIndex)) > 0 Then
                strCopyText = strCopyText & objMwtWin.GetElementSymbol(intIndex) & ": " & strPctCompositions(intIndex) & vbCrLf
            End If
        Next intIndex
    End With
    
    Clipboard.SetText strCopyText, vbCFText

End Sub

Private Sub CopyRTF()
    Clipboard.Clear
    
    rtfFormula(mCurrentFormulaIndex).SelStart = 0
    rtfFormula(mCurrentFormulaIndex).SelLength = Len(rtfFormula(mCurrentFormulaIndex).Text)
    
    Clipboard.SetText RemoveRTFHeightAdjust(rtfFormula(mCurrentFormulaIndex).SelRTF), vbCFRTF
    rtfFormula(mCurrentFormulaIndex).SelLength = 0

End Sub

Public Sub ShowIsoDistributionModeller()

    Dim strFormula As String
    
    strFormula = rtfFormula(mCurrentFormulaIndex).Text
    
    RemoveHeightAdjustChar strFormula
    
    frmIsotopicDistribution.rtfFormula = strFormula
    frmIsotopicDistribution.Show
End Sub

Private Sub EditAbbreviations()
    frmEditAbbrev.ResetValChangedToFalse
    frmEditAbbrev.Show vbModal
    RecalculateAllFormulas
    
    ' Clear lblLoadStatus so it doesn't accumulate messages over time
    frmIntro.lblLoadStatus.Caption = ""
End Sub

Private Sub EditElements()
    frmEditElem.ResetValChangedToFalse
    frmEditElem.Show vbModal
    RecalculateAllFormulas
    
    ' Clear lblLoadStatus so it doesn't accumulate messages over time
    frmIntro.lblLoadStatus.Caption = ""
End Sub

Private Sub EditMenuCheck()
    Dim intIndex As Integer

    If TypeOf Screen.ActiveControl Is RichTextBox Or TypeOf Screen.ActiveControl Is TextBox Then
        mnuCut.Enabled = True
        mnuCopy.Enabled = True
        If Clipboard.GetFormat(vbCFText) Or Clipboard.GetFormat(vbCFRTF) Then
            mnuPaste.Enabled = True
        Else
            mnuPaste.Enabled = False
        End If
        mnuDelete.Enabled = True
        mnuCopyRTF.Enabled = True
    Else
        mnuCut.Enabled = False
        mnuCopy.Enabled = False
        mnuPaste.Enabled = False
        mnuDelete.Enabled = False
        mnuCopyRTF.Enabled = False
    End If
    
    If rtfFormula(mCurrentFormulaIndex).Text = "" Then
        mnuEraseCurrent.Enabled = False
        mnuCopyCurrent.Enabled = False
        mnuExpandAbbrev.Enabled = False
        mnuEmpirical.Enabled = False
    Else
        mnuEraseCurrent.Enabled = True
        mnuCopyCurrent.Enabled = True
        mnuExpandAbbrev.Enabled = True
        mnuEmpirical.Enabled = True
    End If
    
    
    mnuExpandAbbrev.Caption = AppendEllipsesToSingleMenu(mnuExpandAbbrev.Caption, True)
    mnuEmpirical.Caption = AppendEllipsesToSingleMenu(mnuEmpirical.Caption, True)
    mnuEmpirical.Caption = AppendShortcutKeyToSingleMenu(mnuEmpirical.Caption, "Ctrl+E")
    
    If txtMWT(mCurrentFormulaIndex).Visible = False Then
        mnuCopyMWT.Enabled = False
    Else
        mnuCopyMWT.Enabled = True
    End If
    
    For intIndex = 0 To mTopFormulaIndex
        If rtfFormula(intIndex).Text <> "" Then Exit For
    Next intIndex

    If intIndex = mTopFormulaIndex + 1 Then
        mnuEraseAll.Enabled = False
    Else
        mnuEraseAll.Enabled = True
    End If

    If mViewMode = vmdSingleView Then
        mnuCopyPC.Enabled = True
    Else
        mnuCopyPC.Enabled = False
    End If
End Sub

Private Sub FormWideKeyDownHandler(ByRef KeyCode As Integer, ByRef Shift As Integer)

    If mKeyPressAbortPctSolver = 1 Then
        ' Key pressed, so end calculations
        mKeyPressAbortPctSolver = 2
        KeyCode = 0
        Shift = 0
    Else
        If mKeyPressAbortPctSolver = 3 Then
            UpdateGrid grdPC, mCurrentFormulaIndex
        End If
    End If
    
    If (KeyCode = vbKeyF4 And (Shift And vbAltMask)) Then     ' And them in case ctrl or shift was also accidentally pressed
        ' Exit program via exit subroutine
        KeyCode = 0
        Shift = 0
        ExitProgram
    ElseIf KeyCode = vbKeyEscape Then
        If frmProgramPreferences.optExitConfirmation(exmEscapeKeyConfirmExit).value = True Or frmProgramPreferences.optExitConfirmation(exmEscapeKeyDoNotConfirmExit).value = True Then
            ' If Esc key is activated, exit program via exit subroutine
            ExitProgram
        Else
            ' Call ExitProgram subroutine to avoid an error, but don't exit
        End If
    ElseIf KeyCode = vbKeyE And (Shift And vbCtrlMask) Then     ' And them in case alt or shift was also accidentally pressed
        mnuEmpirical_Click
        KeyCode = 0: Shift = 0
    ElseIf KeyCode = vbKeyU And (Shift And vbCtrlMask) Then     ' And them in case alt or shift was also accidentally pressed
        With frmProgramPreferences
            If .chkAutoCopyCurrentMWT = 1 Then
                .chkAutoCopyCurrentMWT = 0
            Else
                .chkAutoCopyCurrentMWT = 1
            End If
        End With
        KeyCode = 0: Shift = 0
    Else
        Select Case KeyCode
        Case vbKeyF1
            ' Windows Help command
        Case vbKeyF2
            ' Copy current formula
        Case vbKeyF3
            ' Change abbreviation recognition mode
            With frmProgramPreferences
                If .optAbbrevType(0).value = True Then
                    .optAbbrevType(1).value = True
                ElseIf .optAbbrevType(1).value = True Then
                    .optAbbrevType(2).value = True
                Else
                    .optAbbrevType(0).value = True
                End If
            End With
            LabelStatus
        Case vbKeyF4
            ' Change Case Mode
            With frmProgramPreferences
                If .optConvertType(0).value = True Then
                    .optConvertType(1).value = True
                ElseIf .optConvertType(1).value = True Then
                    .optConvertType(2).value = True
                Else
                    .optConvertType(0).value = True
                End If
            End With
            LabelStatus
        Case vbKeyF5
            ' Erase all formulas
        Case vbKeyF6
            ' Erase current formula
        Case vbKeyF7
            ' Change caution mode
            frmProgramPreferences.SwapCheck frmProgramPreferences.chkShowCaution
            LabelStatus
        Case vbKeyF8
            ' Switch views
            If mViewMode = vmdMultiView Then
                SetViewMode vmdSingleView
            Else
                SetViewMode vmdMultiView
            End If
        Case vbKeyF9
            frmProgramPreferences.SwapCheck frmProgramPreferences.chkAdvanceOnCalculate
            LabelStatus
        Case vbKeyF10
            ' Reserved by windows for selecting the menus
        Case vbKeyF11
            ' Toggle Percent Solver
            If Not PctSolverParams.SolverOn Then
                TogglePercentComposition 1
            Else
                TogglePercentComposition 0
            End If
            frmMain.SetFocus
            LabelStatus
        Case vbKeyF12
            ' Change standard deviation mode
            With frmProgramPreferences
                If .optStdDevType(0).value = True Then
                    .optStdDevType(1).value = True
                ElseIf .optStdDevType(1).value = True Then
                    .optStdDevType(2).value = True
                ElseIf .optStdDevType(2).value = True Then
                    .optStdDevType(3).value = True
                Else
                    .optStdDevType(0).value = True
                End If
            End With
            LabelStatus
        Case vbKeyReturn
            ' Ignore It
        Case Else
            LabelStatus
        End Select
    End If

End Sub

Public Sub SetFocusToFormulaByIndex(Optional intFormulaIndex As Integer = 0)
    Dim eViewModeSaved As vmdViewModeConstants
    
    ' This is necessary in case a modal form is display when this sub is called
    On Error Resume Next
    
    eViewModeSaved = frmMain.GetViewMode
    If eViewModeSaved = vmdSingleView Then
        SetViewMode vmdMultiView
    End If
    
    mCurrentFormulaIndex = intFormulaIndex
    frmMain.rtfFormula(intFormulaIndex).SetFocus

    If eViewModeSaved = vmdSingleView Then
        SetViewMode vmdSingleView
        frmMain.rtfFormulaSingle.SetFocus
    End If
    
End Sub

Public Sub SetUsingPCGridFalse()
    mUsingPCGrid = False
End Sub

Public Sub SetViewMode(eViewMode As vmdViewModeConstants)
    Dim blnMultiMode As Boolean
    
    If PctSolverParams.SolverOn And eViewMode = vmdMultiView Then
        TogglePercentComposition psmPercentSolverOff
    End If
    
    If eViewMode = vmdMultiView Then blnMultiMode = True
    mViewMode = eViewMode
    
    mnuViewType(vmdMultiView).Checked = blnMultiMode
    mnuViewType(vmdSingleView).Checked = Not blnMultiMode

    fraMulti.Visible = blnMultiMode
    fraSingle.Visible = Not blnMultiMode

    If blnMultiMode Then
        If frmMain.Visible = True Then rtfFormula(mCurrentFormulaIndex).SetFocus
        Calculate True, True, True
    Else
        rtfFormulaSingle.TextRTF = rtfFormula(mCurrentFormulaIndex).TextRTF
        txtMWTSingle.Text = txtMWT(mCurrentFormulaIndex).Text
        lblFormulaSingle.Caption = lblFormula(mCurrentFormulaIndex).Caption
        
        If frmMain.Visible = True Then rtfFormulaSingle.SetFocus
        
        Calculate True, True, True
        If txtMWTSingle.Text = "" Then ClearGrid grdPC
    End If
    
    If cChkBox(frmProgramPreferences.chkShowToolTips) Then
        cmdNewFormula.ToolTipText = ConstructAddFormulaToolTip
    End If

    ResizeFormMain True

End Sub

Public Sub ShowAminoAcidNotationModule()
    frmAminoAcidConverter.Show
End Sub

Public Sub ShowCapillaryFlowModule()
    frmCapillaryCalcs.Show
End Sub

Public Sub ShowFormulaFinder()
    ShowWeightModeWarningDialog True
End Sub

Public Sub ShowMoleMassConverter()
    If rtfFormula(mCurrentFormulaIndex).Text = "" Or objMwtWin.ErrorID <> 0 Then
        frmMMConvert.optWeightSource(1).value = True
    End If
    
    frmMMConvert.Show

End Sub

Public Sub ShowPeptideSequenceModeller()
    ShowWeightModeWarningDialog False
End Sub

Private Sub ShowProgramOptions()
    frmProgramPreferences.lblHiddenFormStatus = "Shown"
    frmProgramPreferences.Show

End Sub

Private Sub ShowWeightModeWarningDialog(blnShow As Boolean)
    ' If blnShow is True then the Formula Finder warning and Formula Finder
    ' window are shown.
    
    ' If blnShow is False then the Fragmentation Modelling window is shown
    
    If objMwtWin.GetElementMode = emAverageMass Then
        ' using average weights, check to warn or not
        If cChkBox(frmProgramPreferences.chkNeverShowFormulaFinderWarning) And _
            Not cChkBox(frmProgramPreferences.chkAlwaysSwitchToIsotopic) Then
            ' Never warn, just leave alone
        ElseIf cChkBox(frmProgramPreferences.chkAlwaysSwitchToIsotopic) Then
            ' Always quick switch to isotopic weights
            SwitchWeightMode emIsotopicMass
        Else   ' Includes case 0 specifically
            ' If not done yet, warn user of weight-type situation, and request future action
            ' First display the proper instructions; must do this before displaying the form modally
            frmFinderModeWarn.DisplayInstructions blnShow
            
            ' Now display the form (modally)
            frmFinderModeWarn.Show vbModal
            
            ' Examine the status of the options
            If frmFinderModeWarn.optWeightChoice(0).value = True Or frmFinderModeWarn.optWeightChoice(1).value = True Then
                ' Quick Switch to isotopic weights now
                SwitchWeightMode emIsotopicMass
            End If
            
            ' Examine the status of the ShowAgain checkbox
            If cChkBox(frmFinderModeWarn.chkShowAgain) Then
                ' Don't show form again
                With frmProgramPreferences
                    If frmFinderModeWarn.optWeightChoice(1).value = True Then
                        ' Switched to isotopic weights this time and should every time
                        .chkAlwaysSwitchToIsotopic = vbChecked  ' This also checks Never show
                    Else
                        ' Stop showing this form and never automatically switch
                        .chkAlwaysSwitchToIsotopic = vbUnchecked
                        .chkNeverShowFormulaFinderWarning = vbChecked
                    End If
                End With
            End If
        End If
    End If

    If blnShow Then
        frmFinder.Show
    Else
        frmFragmentationModelling.Show
    End If

End Sub

Public Sub EraseAllFormulas(Optional boolPromptForErase As Boolean = True)
    Dim eResponse As VbMsgBoxResult
    Dim intIndex As Integer
    
    ' Erase all formulas
    If boolPromptForErase = True Then
        ' Display the dialog box and get user's Response.
        eResponse = YesNoBox(LookupLanguageCaption(3780, "Are you sure you want to erase all the formulas?"), _
                            LookupLanguageCaption(3785, "Erase all Formulas"))
    Else
        eResponse = vbYes
    End If
    
    ' Evaluate the user's Response.
    If eResponse = vbYes Then
        With frmMain
            For intIndex = 0 To frmMain.GetTopFormulaIndex
                .rtfFormula(intIndex).Text = ""
                .txtMWT(intIndex).Text = ""
                .txtMWT(intIndex).Visible = False
            Next intIndex
            .rtfFormulaSingle.Text = ""
            .txtMWTSingle.Text = ""
            .txtMWTSingle.Visible = False
            .SetFocusToFormulaByIndex
            
            Debug.Assert frmMain.GetCurrentFormulaIndex = 0
            
            .SetUsingPCGridFalse
            
            .Calculate True
            .LabelStatus
        End With
    End If

End Sub

Private Sub EraseCurrentFormula()
    ' Erase current formula (F6)
    Dim eResponse As VbMsgBoxResult
    
    ' Display the dialog box and get user's response.
    eResponse = YesNoBox(LookupLanguageCaption(3620, "Are you sure you want to erase the current formula?"), _
                        LookupLanguageCaption(3625, "Erase Current Formula"))

    ' Evaluate the user's response.
    If eResponse = vbYes Then
        rtfFormula(mCurrentFormulaIndex).Text = ""
        
        txtMWT(mCurrentFormulaIndex).Text = ""
        txtMWT(mCurrentFormulaIndex).Visible = False
        
        PctSolverParams.SolverOn = False              ' Turn off % Solver
        mUsingPCGrid = False
        
        Calculate True
    End If
    LabelStatus

End Sub

Private Sub ExpandAbbreviations()
    ' Expand abbreviations of the current formula into their elemental representations
    
    Dim eResponse As VbMsgBoxResult
    
    ' Display the dialog box and get user's response.
    If (frmProgramPreferences.optExitConfirmation(exmEscapeKeyConfirmExit).value = True Or frmProgramPreferences.optExitConfirmation(exmIgnoreEscapeKeyConfirmExit).value = True) Then
        eResponse = YesNoBox(LookupLanguageCaption(3630, "Are you sure you want to expand the abbreviations of the current formula to their elemental equivalents?"), _
                            LookupLanguageCaption(3635, "Expand Abbreviations"))
    Else
        eResponse = vbYes
    End If
    
    ' Evaluate the user's response.
    If eResponse = vbYes Then
        'ExpandAbbreviations
        
        Calculate False, True, True, mCurrentFormulaIndex, False, True
        
        If objMwtWin.ErrorID = 0 Then
            LabelStatus
        Else
            ' Do nothing; error message is being displayed
        End If
    End If

End Sub

Private Sub FormulaKeyDownHandler(ByRef KeyCode As Integer, ByRef Shift As Integer)
   Dim blnErrorPresent As Boolean
    
    If objMwtWin.ErrorID <> 0 Or gBlnErrorPresent Then
        blnErrorPresent = True
        ' Position insertion point on line of error
        If mViewMode = vmdMultiView Then
            rtfFormula(mCurrentFormulaIndex).SetFocus
        Else
            rtfFormulaSingle.SetFocus
        End If
        objMwtWin.ClearError
        gBlnErrorPresent = False: gBlnStatusCaution = False
        mErrorParams.ErrorLine = -1
        mErrorParams.ErrorPosition = -1
        mErrorParams.ErrorCharacter = ""
        LabelStatus
    End If

    If (KeyCode = vbKeyDown Or KeyCode = vbKeyUp) And Not blnErrorPresent Then
        If KeyCode = vbKeyUp Then
            ' Up Arrow
            mCurrentFormulaIndex = mCurrentFormulaIndex - 1
            If mCurrentFormulaIndex < 0 Then mCurrentFormulaIndex = mTopFormulaIndex
        End If
        
        If KeyCode = vbKeyDown Then
           ' Down Arrow
           mCurrentFormulaIndex = mCurrentFormulaIndex + 1
           If mCurrentFormulaIndex > mTopFormulaIndex Then mCurrentFormulaIndex = 0
        End If

        KeyCode = 0
        Shift = 0
        SetFocusToFormulaByIndex mCurrentFormulaIndex
        
        rtfFormulaSingle.Text = rtfFormula(mCurrentFormulaIndex).Text
        lblFormulaSingle.Caption = ConstructFormulaLabel(mCurrentFormulaIndex)
        txtMWTSingle.Text = txtMWT(mCurrentFormulaIndex).Text
        If Len(txtMWTSingle.Text) = 0 Then
            txtMWTSingle.Visible = False
        Else
            txtMWTSingle.Visible = True
        End If
        LabelStatus
        
        If mViewMode = vmdSingleView Then
            Calculate False, True, True
        Else
            Calculate True, True, True
        End If
    End If
    
    Select Case KeyCode
    Case 9
        If (Shift And vbCtrlMask) <> 0 Then
            ' Do not allow Ctrl+Tab
            KeyCode = 0
            Shift = 0
        End If
    Case 33
        SetViewMode vmdSingleView
        If PctSolverParams.SolverOn Then
            ' Page Up
            With grdPC
                .Row = 10
                .SetFocus
                .Tag = True
            End With
            LabelStatus
        End If
    Case 34
        SetViewMode vmdSingleView
        If PctSolverParams.SolverOn Then
            ' Page Down
            With grdPC
                .Row = 0
                .SetFocus
                .Tag = True
            End With
            LabelStatus
        End If
    End Select
End Sub

Private Sub FormulaKeyPressHandler(ThisRTFTextBox As RichTextBox, KeyAscii As Integer)
    Dim intSelStartSaved As Integer
    
    If gBlnStatusCaution And objMwtWin.ErrorID = 0 Then
        gBlnStatusCaution = False
        If KeyAscii <> 0 Then LabelStatus
    End If

    If PctSolverParams.SolverOn And mDifferencesFormShown Then
        With PctSolverParams
            mDifferencesFormShown = False
            .SolverWorking = True
    
            Debug.Assert gBlnShowStdDevWithMass = .StdDevModeEnabledSaved
            
            .XVal = .BestXVal
            rtfFormula(mCurrentFormulaIndex).Tag = "Changed"
            intSelStartSaved = rtfFormulaSingle.SelStart
            
            Calculate False, True, True, mCurrentFormulaIndex, False, False, True, .XVal
            
            rtfFormulaSingle.SelStart = intSelStartSaved
            .SolverWorking = False
        End With
    End If

    ' These are special-case keypress events that are not handled by RTFBoxKeyPressHandler
    Select Case KeyAscii
    Case 40 To 41, 43, 45, 48 To 57, 62, 65 To 90, 91, 93, 94, 95, 97 To 122, 123, 125
        ' Valid Characters (see RTFBoxKeyPressHandler for details)
        rtfFormula(mCurrentFormulaIndex).Tag = FORMULA_CHANGED
    Case 8
        ' Backspace; it is valid
        rtfFormula(mCurrentFormulaIndex).Tag = FORMULA_CHANGED
    Case vbKeyReturn
        InitiateCalculate
        KeyAscii = 0
    End Select
    
    ' Check the validity of the key using RTFBoxKeyPressHandler
    If KeyAscii <> 0 Then RTFBoxKeyPressHandler Me, ThisRTFTextBox, KeyAscii, True

End Sub

Public Function GetCurrentFormulaIndex() As Integer
    GetCurrentFormulaIndex = mCurrentFormulaIndex
End Function

Public Function GetTopFormulaIndex() As Integer
    GetTopFormulaIndex = mTopFormulaIndex
End Function

Public Function GetViewMode() As vmdViewModeConstants
    GetViewMode = mViewMode
End Function

Private Sub InitiateCalculate()
    Dim lngErrorID As Long
    
    If Not PctSolverParams.SolverOn Then
        ' Change mouse pointer to hourglass
        Me.MousePointer = vbHourglass
        
        lngErrorID = Calculate(True)
        
        ' Change mouse pointer to default
        Me.MousePointer = vbDefault
        
        If Len(rtfFormula(mCurrentFormulaIndex).Text) = 0 Then
            If mViewMode = vmdMultiView Then
                txtMWT(mCurrentFormulaIndex).Visible = False
                txtMWT(mCurrentFormulaIndex).Text = ""
            Else
                txtMWTSingle.Visible = False
                txtMWTSingle.Text = ""
            End If
        End If
    Else
        InitiatePercentSolver
        
        'The following is necessary to prevent grdPC from being updated just after .InitiatePercentSolver() finishes
        mIgnoreNextGotFocus = True

    End If
            
    ' Update the status line
    LabelStatus lngErrorID

End Sub

Private Sub InitializeFormWideVariables()
    
    ' Set default Public variable values
    mTopFormulaIndex = 0                ' Number of formulas fields displayed
    mCurrentFormulaIndex = 0            ' Current Formula number
        
    mCurrentFormulaIndex = 0            ' Current Forumla

    mPercentCompositionGridRowCount = 11

End Sub

Private Sub InitiatePercentSolver()
    ' Solve for the optimum value for x in a formula such that the
    ' computed percent composition values match those entered by the user
    
    Const REPEAT_COUNT_STOP_VAL = 3
    
    Dim intXCharLoc As Integer
    Dim blnXPresent As Boolean
    Dim lngErrorID As Long
    
    Dim dblXValMinimum As Double, dblXValMaximum As Double
    Dim dblXValPreviousMax As Double, dblStepSize As Double
    Dim dblWork As Double
    Dim intRepeatCount As Integer, intIndex As Integer
    Dim intColonLoc As Integer
    
    If PctSolverParams.SolverWorking Then Exit Sub
    
    ' Make sure an x is present after a bracket in the formula
    intXCharLoc = InStr(rtfFormula(mCurrentFormulaIndex).Text, "[")
    If intXCharLoc > 0 And LCase(Mid(rtfFormula(mCurrentFormulaIndex).Text, intXCharLoc + 1, 1)) = "x" Then
        blnXPresent = True
    End If

    If Not blnXPresent Or mPctCompValues(mCurrentFormulaIndex).LockedValuesCount = 0 Then
        ' Either x is not present in the formula after a bracket or no values locked
        ' Calculate normally
        mKeyPressAbortPctSolver = 2
        Calculate True
    Else
        ' Percent Solver Routine
        dblXValMinimum = 0              ' Current minimum x value
        dblXValMaximum = 1000           ' Current maximum x value
        mKeyPressAbortPctSolver = 1
        intRepeatCount = 0

        With PctSolverParams
            .SolverWorking = True
            .DivisionCount = PCT_SOLVER_INITIAL_DIVISION_COUNT
            .blnMaxReached = False
            .StdDevModeEnabledSaved = gBlnShowStdDevWithMass
            .BestXSumResiduals = 10 ^ 30                       ' Set .BestXSumResiduals to a really big number
        End With

        lblStatus.ForeColor = QBColor(COLOR_CALC)
        lblStatus.Caption = "  " & LookupLanguageCaption(3700, "Calculating, press any key or click the mouse to stop.")

        ' Change mouse pointer to hourglass
        Me.MousePointer = vbHourglass
        Do
            dblXValPreviousMax = dblXValMaximum
            PctSolverParams.SolverCounter = -1
            dblStepSize = ((dblXValMaximum - dblXValMinimum) / PctSolverParams.DivisionCount)
            If dblStepSize = 0 Then dblStepSize = -1
            PctSolverParams.IterationStopValue = PctSolverParams.DivisionCount + 1
            
            For dblWork = dblXValMinimum To dblXValMaximum Step dblStepSize
                DoEvents
                If mKeyPressAbortPctSolver <> 1 Then Exit Do
                PctSolverParams.SolverCounter = PctSolverParams.SolverCounter + 1
                PctSolverParams.XVal = dblWork
                
                If PctSolverParams.SolverCounter = PctSolverParams.DivisionCount / 2 Then
                    ' Update the current x value half way through the iterations
                    lngErrorID = Calculate(False, True, True, mCurrentFormulaIndex, False, False, True, PctSolverParams.XVal)
                Else
                    lngErrorID = Calculate(False, False, True, mCurrentFormulaIndex, False, False, True, PctSolverParams.XVal)
                End If
                If lngErrorID <> 0 Then Exit Do
                If PctSolverParams.SolverCounter > PctSolverParams.IterationStopValue Then Exit For
            Next dblWork

            If Not PctSolverParams.blnMaxReached And PctSolverParams.BestXVal = dblXValMaximum Then
                ' The best X Val was dblXValMaximum, and we haven't yet found the maximum,
                '  so set the step size to the previous maximum (essentially zooming out the search)
                dblStepSize = dblXValPreviousMax
            Else
                ' The best X val was less than dblXValMaximum, so set blnMaxReached to True
                ' The best X val must be within dblStepSize of .BestXVal
                PctSolverParams.blnMaxReached = True
            End If
            
            dblXValMinimum = PctSolverParams.BestXVal - dblStepSize
            dblXValMaximum = PctSolverParams.BestXVal + dblStepSize
            If dblXValMinimum < 0 Then dblXValMinimum = 0
            If dblXValMaximum > 1E+19 Then mKeyPressAbortPctSolver = 0
            
            If Format(dblXValMaximum, "0.0000000E+00") = Format(dblXValPreviousMax, "0.0000000E+00") Then
                intRepeatCount = intRepeatCount + 1
                With PctSolverParams
                    ' If a value is repeated, then test 10x, 50x, or 100 times more values on the next iteration
                    Select Case intRepeatCount
                    Case 1
                        If .DivisionCount < PCT_SOLVER_INITIAL_DIVISION_COUNT * 10 Then .DivisionCount = PCT_SOLVER_INITIAL_DIVISION_COUNT * 10
                    Case 2
                        If .DivisionCount < PCT_SOLVER_INITIAL_DIVISION_COUNT * 50 Then .DivisionCount = PCT_SOLVER_INITIAL_DIVISION_COUNT * 50
                    Case 3
                        If .DivisionCount < PCT_SOLVER_INITIAL_DIVISION_COUNT * 100 Then .DivisionCount = PCT_SOLVER_INITIAL_DIVISION_COUNT * 100
                    Case Else
                        ' Leave .DivisionCount unchanged
                    End Select
                End With
            End If
            If intRepeatCount = REPEAT_COUNT_STOP_VAL Or dblStepSize = 0 Then mKeyPressAbortPctSolver = 0
        Loop
        mKeyPressAbortPctSolver = 2

        ' Change mouse pointer to default
        MousePointer = vbDefault

        If lngErrorID = 0 Then
            PctSolverParams.XVal = PctSolverParams.BestXVal
            rtfFormula(mCurrentFormulaIndex).Tag = "Changed"
            lngErrorID = Calculate(False, True, True, mCurrentFormulaIndex, False, False, True, PctSolverParams.XVal)
            Debug.Assert lngErrorID = 0
            
            lblValueForX.ForeColor = QBColor(COLOR_SOLVER)
            lblValueForX.Caption = LookupLanguageCaption(3710, "x is") & " " & Format(PctSolverParams.BestXVal, "#0.0######")

            With frmDiff
                .lblDiff1.Caption = ""
                .lblDiff2.Caption = ""
                .lblDiff3.Caption = ""
            End With

            ' Print differences
            AddToDiffForm LookupLanguageCaption(3720, "Calculated Value"), 1
            AddToDiffForm LookupLanguageCaption(3730, "Target"), 2
            AddToDiffForm LookupLanguageCaption(3740, "Difference from Target"), 3

             For intIndex = 1 To mPctCompValues(mCurrentFormulaIndex).TotalElements
                 With mPctCompValues(mCurrentFormulaIndex).GridCell(intIndex)
                    ' Check Locking Status
                    If .Locked Then
                       ' The element is locked
                       intColonLoc = InStr(.Text, ":")
                       AddToDiffForm Left(.Text, intColonLoc) & " " & Format(.value, "##0.0######"), 1
                       AddToDiffForm CStr(.Goal), 2
                       AddToDiffForm Format(.value - .Goal, "##0.0######"), 3
                    End If
                End With
             Next intIndex
            If frmDiff.WindowState = vbMinimized Then frmDiff.WindowState = vbNormal
            frmDiff.Show
            mDifferencesFormShown = True

            LabelStatus
        Else
            ' Error present. Error position was not highlighted because solver was equal to 1
            If rtfFormulaSingle.Tag = "Highlighted" Then
                PctSolverParams.SolverWorking = False
                rtfFormulaSingle.Tag = ""
                UpdateAndFormatFormulaSingle
            End If
        End If
        
        gBlnShowStdDevWithMass = PctSolverParams.StdDevModeEnabledSaved
        PctSolverParams.SolverWorking = False
    End If

End Sub


Public Sub LabelStatus(Optional lngLocallyCaughtErrorID As Long = 0)
    
    If objMwtWin.ErrorID = 0 And lngLocallyCaughtErrorID = 0 And Not gBlnStatusCaution And Not gBlnErrorPresent Then
        If mUsingPCGrid = False Then
            If PctSolverParams.SolverOn Then
                lblStatus.ForeColor = QBColor(COLOR_DIREC)
                lblStatus.Caption = LookupLanguageCaption(3660, "Use Page Up/Down or Up/Down arrows to move to the percents (F11 exits Percent Solver mode).")
            End If
        Else
            If PctSolverParams.SolverOn Then
                ' Percent Solver is On
                lblStatus.ForeColor = QBColor(COLOR_DIREC)
                lblStatus.Caption = LookupLanguageCaption(3665, "Press Enter or Click to change a percentage (F11 exits Percent Solver Mode).")
            End If
        End If
    End If
    
    If gBlnStatusCaution = True Then
        gBlnStatusCaution = False
    Else
        If Not PctSolverParams.SolverOn And objMwtWin.ErrorID = 0 And lngLocallyCaughtErrorID = 0 Then
            lblStatus.ForeColor = vbWindowText
            lblStatus.Caption = LookupLanguageCaption(3670, "Ready")
            Select Case objMwtWin.GetElementMode()
            Case emIsotopicMass
                lblStatus.Caption = lblStatus.Caption & " " & LookupLanguageCaption(3830, "(using isotopic elemental weights)")
            Case emIntegerMass
                lblStatus.Caption = lblStatus.Caption & " " & LookupLanguageCaption(3840, "(using integer isotopic weights)")
            Case Else   ' The default, using average atomic weights
                        ' Don't append anything; just show Ready
            End Select
        End If
    End If
    
End Sub

Private Function ParseCommandLine(blnPreParseCommandLine As Boolean) As Boolean
    ' Parses the command line
    ' Returns true if /?, /Help, or an invalid argument was used at the command line
    ' If blnPreParseCommandLine = True, then only checks for /X, /Z, or /D and does not
    '  perform batch analysis
    
    ' Declare local variables
    Dim strArguments(15) As String
    Dim strMessage As String, strWork As String, strCommandLine As String
    Dim intMaxArgs As Integer, intNumArgs As Integer, intStatus As Integer, intIndex As Integer
    Dim boolInAnArgument As Boolean

    strCommandLine = Command()
    intNumArgs = 0
    If strCommandLine <> "" Then
        intMaxArgs = 15
        For intIndex = 1 To Len(strCommandLine)
            strWork = Mid(strCommandLine, intIndex, 1)
            ' Test for character being a blank or a tab.
            If strWork <> " " And strWork <> vbTab Then
                ' Neither blank nor tab. Test if we're already inside an argument.
                If Not boolInAnArgument Then
                    ' Found the start of a new argument.
                    ' Test for too many arguments.
                    If intNumArgs = intMaxArgs Then Exit For
                    intNumArgs = intNumArgs + 1
                    boolInAnArgument = True
                End If
                ' Add the character to the current argument.
                strArguments(intNumArgs) = strArguments(intNumArgs) + strWork
            Else
                ' Found a blank or a tab.
                ' Set "Not in an argument" flag to FALSE.
                boolInAnArgument = False
            End If
        Next intIndex
    
        For intIndex = 1 To intNumArgs
            strWork = strArguments(intIndex)
            If Left(strWork, 1) = "/" Or Left(strWork, 1) = "-" Then strWork = Mid(strWork, 2)
    
            intStatus = 0
            Select Case UCase(Left(strWork, 1))
            Case "F"
                mBlnBatchProcess = True
                If Mid(strWork, 2, 1) = ":" Then
                    mBatchProcessFilename = Mid(strWork, 3)
                Else
                    mBatchProcessFilename = Mid(strWork, 2)
                End If
            Case "O"
                If Mid(strWork, 2, 1) = ":" Then
                    mBatchProcessOutfile = Mid(strWork, 3)
                Else
                    mBatchProcessOutfile = Mid(strWork, 2)
                End If
            Case "Y"
                mBatchProcessOverwrite = True
            Case "X"
                gBlnWriteFilesOnDrive = False
            Case "Z"
                gBlnAccessFilesOnDrive = False
                gBlnWriteFilesOnDrive = False
            Case "D"
                If UCase(strWork) = "DEBUG" Then
                    ' Ignore it; user is debugging software
                Else
                    intStatus = 2
                End If
            Case "?", "h", "H"
                If Not blnPreParseCommandLine Then intStatus = 2
            Case Else
                If Not blnPreParseCommandLine Then intStatus = 1
            End Select
            
            If intStatus <> 0 Then
                strMessage = vbCrLf
                If (Left(strWork, 1) = "?" Or Left(strWork, 1) = "H") And intStatus = 2 Then
                    strMessage = strMessage & vbCrLf & "Command line options."
                Else
                    strMessage = strMessage & vbCrLf & "Incorrect options for command line:  " & strCommandLine
                    strMessage = strMessage + vbCrLf
                End If
                strMessage = strMessage & vbCrLf & "Syntax:  MWTWIN [/F:filename] [/O:outfile] [/Y] [/X] [/?]"
                strMessage = strMessage + vbCrLf
                strMessage = strMessage & vbCrLf & "Available options are:"
                strMessage = strMessage & vbCrLf & "/F:filename  to start MWTWIN in batch mode, calculating weights for compounds in filename, writing"
                strMessage = strMessage & vbCrLf & "                    results to filename.out, and exiting"
                strMessage = strMessage & vbCrLf & "/O:outfile     to specify the name of the output file for batch mode"
                strMessage = strMessage & vbCrLf & "/Y               to overwrite outfile without prompting the user"
                strMessage = strMessage & vbCrLf & "/X               to prevent MWTWIN from attempting to write abbreviation, element, or default value files"
                strMessage = strMessage & vbCrLf & "/Z               to prevent MWTWIN from attempting to read or write abbreviations/elements/values files"
                '
                ' Secret option: /Debug     to show debug prompts
                '
                strMessage = strMessage & vbCrLf & "/? or /Help  to display this screen."
                strMessage = strMessage + vbCrLf
                strMessage = strMessage & vbCrLf & "See MWTWIN.CHM for more help."
                strMessage = strMessage & vbCrLf & "Program by Matthew Monroe.  Send E-Mail to Matt@alchemistmatt.com or AlchemistMatt@Yahoo.com"
                strMessage = strMessage & vbCrLf & "WWW - http://www.alchemistmatt.com/ and http://www.come.to/alchemistmatt"
                strMessage = strMessage + vbCrLf
                strMessage = strMessage & vbCrLf & "This program is Freeware; distribute freely."
                
                frmIntro.cmdExit.Visible = True
                frmIntro.cmdOK.Visible = False
                frmIntro.cmdExit.Default = False
                frmIntro.cmdExit.Cancel = True

                frmIntro.lblLoadStatus.Visible = True
                frmIntro.lblLoadStatus.Caption = strMessage
                frmIntro.MousePointer = vbDefault
            End If
        Next intIndex
    End If

    If intStatus <> 0 Then
        ParseCommandLine = True
    Else
        ParseCommandLine = False
    End If

End Function

Private Sub PercentCompGridClickHandler()
    ' The grid was clicked
    Const DISPLAY_PRECISION = 5
    
    Dim strMessage As String
    Dim intDereferencedIndex As Integer, strDefaultValue As String
    
    If mKeyPressAbortPctSolver = 1 Then
        ' Key pressed, so end calculations
        mKeyPressAbortPctSolver = 2
    End If

    If cChkBox(frmProgramPreferences.chkBracketsAsParentheses) Then
        MsgBox LookupMessage(550), vbOKOnly + vbInformation, LookupMessage(555)
        Exit Sub
    End If
    
    If grdPC.Text <> "" Then
        If grdPC.Text <> "" And Not PctSolverParams.SolverOn Then
            TogglePercentComposition psmPercentSolverOn
        End If
        
        ' The message of the dialog box.
        strMessage = LookupLanguageCaption(3650, "Set the Percent Solver target percentage of this element to a percentage.  Select Reset to un-target the percentage or Cancel to ignore any changes.")
        
        ' Display the dialog box and get user's response.
        With frmChangeValue
            .cmdReset.Caption = LookupLanguageCaption(9220, "&Reset")
            .lblHiddenButtonClickStatus = BUTTON_NOT_CLICKED_YET
            .lblInstructions.Caption = strMessage
            .rtfValue.Visible = False
            .txtValue.Visible = True
        End With

        intDereferencedIndex = (grdPC.Row + 1) + mPercentCompositionGridRowCount * (grdPC.Col)
        ' Display the dialog box and get user's response.
        
        With mPctCompValues(mCurrentFormulaIndex).GridCell(intDereferencedIndex)
            If .Locked Then
                strDefaultValue = Trim(CStr(.Goal))
            Else
                strDefaultValue = Trim(CStr(Round(.value, DISPLAY_PRECISION)))
            End If
        End With
        frmChangeValue.txtValue = strDefaultValue
        
        frmChangeValue.Show vbModal
        
        If frmChangeValue.lblHiddenButtonClickStatus = BUTTON_NOT_CLICKED_YET Then frmChangeValue.lblHiddenButtonClickStatus = BUTTON_CANCEL
        
        If Not frmChangeValue.lblHiddenButtonClickStatus = BUTTON_CANCEL Then
            With mPctCompValues(mCurrentFormulaIndex)
                If frmChangeValue.lblHiddenButtonClickStatus = BUTTON_RESET Then
                    If .GridCell(intDereferencedIndex).Locked Then
                        .GridCell(intDereferencedIndex).Locked = False
                        .LockedValuesCount = .LockedValuesCount - 1
                    End If
                    .GridCell(intDereferencedIndex).Goal = 0
                Else
                    ' Lock intDereferencedIndex
                    If Trim(frmChangeValue.txtValue) <> Trim(CStr(Round(.GridCell(intDereferencedIndex).value, DISPLAY_PRECISION))) And frmChangeValue.txtValue <> "" Then
                        If Not .GridCell(intDereferencedIndex).Locked Then
                            .GridCell(intDereferencedIndex).Locked = True
                            .LockedValuesCount = .LockedValuesCount + 1
                        End If
                        .GridCell(intDereferencedIndex).Goal = CDblSafe(frmChangeValue.txtValue)
                    End If
                End If
            End With
            UpdateGrid grdPC, mCurrentFormulaIndex
        End If
    End If
End Sub

Private Function PerformPercentSolverCalculations(intFormulaIndex As Integer) As Long
    ' Returns 0 if success; Error number if error
    
    Dim dblSumResidualsSquared As Double
    Dim blnAtLeastOneLockedValue As Boolean
    Dim intPCompLineIndex As Integer
    Dim lngErrorID As Long
    
    ' % Solver Calculations
    dblSumResidualsSquared = 0
    blnAtLeastOneLockedValue = False
    For intPCompLineIndex = 1 To mPctCompValues(intFormulaIndex).TotalElements
        With mPctCompValues(intFormulaIndex).GridCell(intPCompLineIndex)
            If .Locked Then
                ' Percent has been locked by user
                blnAtLeastOneLockedValue = True
                If .Goal > 100 Then
                    lngErrorID = 50
                    ' Set intPCompLineIndex = max value so for loop gets exitted
                    ' Cannot use an Exit For statement because inside a With statement
                    intPCompLineIndex = mPctCompValues(intFormulaIndex).TotalElements
                Else
                    ' dblSumResidualsSquared is a sum of the residuals squared
                    ' The fit is minimized when dblSumResidualsSquared is minimized
                    dblSumResidualsSquared = dblSumResidualsSquared + (.value - .Goal) ^ 2
                End If
            End If
        End With
    Next intPCompLineIndex

    If blnAtLeastOneLockedValue Then
        If dblSumResidualsSquared <= PctSolverParams.BestXSumResiduals Then
            PctSolverParams.BestXVal = PctSolverParams.XVal         ' The x value to give this best result
            PctSolverParams.BestXSumResiduals = dblSumResidualsSquared     ' The best result for the given x value
        End If
    End If

    If lngErrorID <> 0 Then          ' Inform user of error and location
        RaiseLabelStatusError lngErrorID
        ResizeFormMain
    End If

    PerformPercentSolverCalculations = lngErrorID
End Function

Private Sub PositionFormControls()
    ' Position Objects
    lblStatus.Left = 0          ' Toolbar, remaining locs set in Resize Procedure
    
    With fraMulti
        .Left = 120         ' Multi Frame
        .Top = 50
        .Height = 1025
    End With
    With lblFormula(0)
        .Left = 120   ' Formula Label
        .Top = 350
    End With
    With rtfFormula(0)
        .Left = 1080   ' Formula 1 text box
        .Top = 240
        .ZOrder 0
    End With
    With txtMWT(0)
        .Left = 1080       ' MW= Label
        .Top = 725
    End With
    With fraSingle
        .Left = 120        ' Single Frame
        .Top = 50
        .Height = 4200
    End With
    With lblFormulaSingle
        .Left = 120     ' Formula Label
        .Top = lblFormula(0).Top
    End With
    With rtfFormulaSingle
        .Left = 1080    ' Formula text box
        .Top = 240
    End With
    With txtMWTSingle
        .Left = 1080        ' MW Label
        .Top = 725
    End With
    With grdPC
        .Top = lblFormulaSingle.Top + 725
        .Left = lblFormulaSingle.Left
        '.HighLight = False
    End With
    cmdCalculate.Top = 150          ' Calculate button
    cmdNewFormula.Top = 600         ' NewFormula button
    
    With lblValueForX
        .Top = 1050
        .Visible = False
        .Caption = ""
    End With

    mnuRightClick.Visible = False
    
End Sub

Private Sub PrintResults()
    Dim intIndex As Integer
    
    With CommonDialog1
        .CancelError = True
        .flags = &H8 + &H4 + &H100000  'cdlPDNoPageNums + cdlPDNoSelection + cdlPDHidePrintToFile
        On Error Resume Next
        .ShowPrinter
    End With
    
    If Err.Number <> 0 Then
        ' No file selected from the Open File dialog box.
        Exit Sub
    End If

    If mViewMode = vmdMultiView Then
        ' Multi Mode
        Printer.Print LookupLanguageCaption(3760, "Molecular Weight Calculator Results")
        Printer.Print ""
        For intIndex = 0 To mTopFormulaIndex
             Printer.Print lblFormula(intIndex + 1).Caption; Tab(15); rtfFormula(intIndex).Text
             Printer.Print Tab(15); txtMWT(intIndex).Text
             Printer.Print ""
        Next intIndex
    Else
        ' Single Mode
        Printer.Print lblFormula(mCurrentFormulaIndex + 1).Caption; ":"; Tab(15); rtfFormulaSingle.Text
        Printer.Print Tab(15); txtMWTSingle.Text
        Printer.Print ""
        For intIndex = 1 To mPctCompValues(mCurrentFormulaIndex).TotalElements
              Printer.Print mPctCompValues(mCurrentFormulaIndex).GridCell(intIndex).Text
        Next intIndex
    End If

    Printer.EndDoc
    
End Sub

Public Sub ResizeFormMain(Optional blnEnlargeToMinimums As Boolean = False, Optional blnUpdateGridPC As Boolean = True)
    Dim intIndex As Integer, minWidth As Single, minHeight As Single
    Dim lngPreferredTop As Long, lngMinimumTop As Long, lngCompareVal As Long
    Dim blnAllowingShrinking As Boolean, blnWidthsAdjusted As Boolean
    Dim lngDesiredValue As Long
    
On Error GoTo ResizeFormMainErrorHandler

    If Me.WindowState = vbMinimized Then Exit Sub

    Me.ScaleMode = vbTwips
    blnAllowingShrinking = True
    
    With frmMain
        If .WindowState = vbNormal Then
            If mnuStayOnTop.Tag = FORM_STAY_ON_TOP_NEWLY_ENABLED And minWidth < mCurrentFormWidth Then
                minWidth = mCurrentFormWidth
            End If
            If mViewMode = vmdMultiView Then
                If minWidth < 5000 Then minWidth = 5000
            Else
                If minWidth < 6300 Then minWidth = 6300
            End If
            
            If .Width < minWidth Then
                If blnEnlargeToMinimums Then
                    .Width = minWidth
                Else
                    blnAllowingShrinking = False
                End If
            End If
            
            If mViewMode = vmdMultiView Then
                minHeight = MAIN_HEIGHT + mTopFormulaIndex * FORMULA_OFFSET_AMOUNT - 1500
            Else
                minHeight = MAIN_HEIGHT_SINGLE
            End If
            
            If frmProgramPreferences.lblHiddenDefaultsLoadedStatus = "DefaultsLoaded" Then
                If minHeight < MAIN_HEIGHT And cChkBox(frmProgramPreferences.chkShowQuickSwitch) Then
                    minHeight = MAIN_HEIGHT
                End If
            End If
            
            If mnuStayOnTop.Tag = FORM_STAY_ON_TOP_NEWLY_ENABLED And minHeight < mCurrentFormHeight Then
                minHeight = mCurrentFormHeight
            End If
            
            If .Height < minHeight Then
                If blnEnlargeToMinimums Then
                    .Height = minHeight
                End If
            End If
            
            If .Width > Screen.Width Then
                .Width = Screen.Width
            End If
        End If
    End With
    mnuStayOnTop.Tag = ""
    
    If frmMain.WindowState <> vbMinimized Then
        lngDesiredValue = frmMain.ScaleWidth - cmdCalculate.Width - 500
        If mViewMode = vmdMultiView Then
            ' Position Objects in Multi View
            If blnAllowingShrinking Or fraMulti.Width < lngDesiredValue Then
                blnWidthsAdjusted = True
                fraMulti.Width = lngDesiredValue
                For intIndex = 0 To mTopFormulaIndex
                    rtfFormula(intIndex).Width = fraMulti.Width - 1200     ' Formula intIndex text box
                    txtMWT(intIndex).Width = rtfFormula(intIndex).Width
                Next intIndex
            End If
        Else
            ' Position Objects in Single View
            If blnAllowingShrinking Or fraSingle.Width < lngDesiredValue Then
                blnWidthsAdjusted = True
                fraSingle.Width = lngDesiredValue
                rtfFormulaSingle.Width = fraSingle.Width - 1200     ' FormulaSingle text box
                txtMWTSingle.Width = rtfFormulaSingle.Width
            End If
        End If
        
        If blnUpdateGridPC Then
            UpdateGrid grdPC, mCurrentFormulaIndex
        End If
        
        lngPreferredTop = frmMain.ScaleHeight - lblStatus.Height
        If mViewMode = vmdMultiView Then
            lngMinimumTop = fraMulti.Top + fraMulti.Height + 60
            If optElementMode(2).Visible Then
                lngCompareVal = optElementMode(2).Top + optElementMode(2).Height
                If lngMinimumTop < lngCompareVal Then
                    lngMinimumTop = lngCompareVal
                End If
            End If
        Else
            lngMinimumTop = fraSingle.Top + fraSingle.Height + 60
        End If
        
        If lngPreferredTop < lngMinimumTop Then lngPreferredTop = lngMinimumTop
        lblStatus.Top = lngPreferredTop
        If lblStatus.Width < minWidth Then
            lblStatus.Width = minWidth
        Else
            lblStatus.Width = Me.ScaleWidth
        End If
        
        If blnAllowingShrinking Or blnWidthsAdjusted Then
            With frmMain
                lblStatus.Width = .ScaleWidth         ' Status Bar
                lblStatus.Width = .ScaleWidth         ' Status Bar
                cmdCalculate.Left = .Width - cmdCalculate.Width - 350     ' Calculate Button
                cmdNewFormula.Left = cmdCalculate.Left   ' New Formula Button
                lblValueForX.Left = cmdCalculate.Left
                lblQuickSwitch.Left = .Width - cmdCalculate.Width - 250
            End With
            
            For intIndex = 0 To 2
                optElementMode(intIndex).Left = lblQuickSwitch.Left
            Next intIndex
        End If
    End If

    Exit Sub

ResizeFormMainErrorHandler:
    Debug.Assert False
    GeneralErrorHandler "frmMain|ResizeFormMain", Err.Number, Err.Description
    
End Sub

Private Sub RestoreDefaultValues()
    Dim eResponse As VbMsgBoxResult
            
    ' Display the dialog box and get user's response.
    eResponse = MsgBox(LookupLanguageCaption(3640, "Restoring the default values and formulas will clear the current formulas.  Are you sure you want to do this?"), vbYesNo + vbQuestion + vbDefaultButton2, _
                      LookupLanguageCaption(3645, "Restoring Values and Formulas"))

    If eResponse = vbYes Then
        ' Change cursor to be on formula 1
        SetFocusToFormulaByIndex 0
        
        EraseAllFormulas False
        SwitchWeightMode emAverageMass
        
        LoadValuesAndFormulas True

        lblStatus.ForeColor = vbWindowText
        lblStatus.Caption = LookupLanguageCaption(3770, "Default values and formulas restored.")
        Calculate True
    End If

End Sub

Private Sub RTFHighlightCharacter(intWorkLine As Integer, intWorkCharLoc As Integer)
    
    Dim strOriginalText As String, strNewText As String
    
    If mViewMode = vmdMultiView Then
        strOriginalText = rtfFormula(intWorkLine).Text
        
        ' Mark the error place with a % sign
        strNewText = Left(strOriginalText, intWorkCharLoc - 1) & "%" & Mid(strOriginalText, intWorkCharLoc)
        If Right(strNewText, 1) = " " Then strNewText = Left(strNewText, Len(strNewText) - 1)
        rtfFormula(intWorkLine).Text = strNewText
        rtfFormula(intWorkLine).Tag = FORMULA_HIGHLIGHTED
    Else
        strOriginalText = rtfFormulaSingle.Text
        
        ' Mark the error place with a % sign
        strNewText = Left(strOriginalText, intWorkCharLoc - 1) & "%" & Mid(strOriginalText, intWorkCharLoc)
        If Right(strNewText, 1) = " " Then strNewText = Left(strNewText, Len(strNewText) - 1)
        rtfFormulaSingle.Text = strNewText
        rtfFormulaSingle.Tag = FORMULA_HIGHLIGHTED
    End If
End Sub

Public Sub ShowHideQuickSwitch(blnShowControl As Boolean)
    Dim intIndex As Integer, eElementWeightType As emElementModeConstants
    
    lblQuickSwitch.Visible = blnShowControl
    
    For intIndex = 0 To 2
        optElementMode(intIndex).Visible = blnShowControl
    Next intIndex

    eElementWeightType = objMwtWin.GetElementMode()
    
    If eElementWeightType > 0 And _
       (optElementMode(0).value = True And eElementWeightType <> emAverageMass Or _
        optElementMode(1).value = True And eElementWeightType <> emIsotopicMass Or _
        optElementMode(2).value = True And eElementWeightType <> emIntegerMass) Then
            optElementMode(eElementWeightType - 1).value = True
    End If
    
End Sub

Private Sub TogglePercentComposition(ePercentSolverMode As psmPercentSolverModeConstants)
    Dim intIndex As Integer
    
    If mnuPercentType(psmPercentSolverOff).Checked = True And _
       ePercentSolverMode = psmPercentSolverOn Then
        ' Turn Percent Solver On (if not on yet)
        If Not cChkBox(frmProgramPreferences.chkBracketsAsParentheses) Then
            PctSolverParams.SolverOn = True
            
            PctSolverParams.XVal = 1
            PctSolverParams.BestXVal = 1
            
            SetViewMode vmdSingleView
            lblValueForX.Visible = True
            lblValueForX.ForeColor = vbWindowText
            lblValueForX.Caption = LookupLanguageCaption(3750, "% Solver On")
            For intIndex = 0 To 3
                If frmProgramPreferences.optStdDevType(intIndex).value = True Then
                    mStdDevModSaved = intIndex
                    Exit For
                End If
            Next intIndex
            
            If mStdDevModSaved = 1 Or mStdDevModSaved = 2 Then
                frmProgramPreferences.optStdDevType(0).value = True
            End If
            
            cmdNewFormula.Enabled = False
        Else
            MsgBox LookupMessage(550), vbOKOnly + vbInformation, LookupMessage(555)
        End If
    Else
        If ePercentSolverMode = psmPercentSolverOff Then
            ' Turn Percent Solver Off
            PctSolverParams.SolverOn = False
            
            lblValueForX.Visible = False
            lblValueForX.Caption = ""
            frmProgramPreferences.optStdDevType(mStdDevModSaved).value = True
        End If
        cmdNewFormula.Enabled = True
    End If
    
    mnuPercentType(psmPercentSolverOff).Checked = Not PctSolverParams.SolverOn
    mnuPercentType(psmPercentSolverOn).Checked = PctSolverParams.SolverOn

    LabelStatus

End Sub

Private Sub ToggleStayOnTop()
    
    ' Add or remove the check mark from the menu.
    mnuStayOnTop.Checked = Not mnuStayOnTop.Checked
   
    With frmMain
        mCurrentFormWidth = .Width
        mCurrentFormHeight = .Height
    End With
    
    mnuStayOnTop.Tag = FORM_STAY_ON_TOP_NEWLY_ENABLED
    
    Me.ScaleMode = vbTwips
    
    WindowStayOnTop Me.hwnd, mnuStayOnTop.Checked, Me.ScaleX(Me.Left, vbTwips, vbPixels), Me.ScaleY(Me.Top, vbTwips, vbPixels), Me.ScaleX(Me.Width, vbTwips, vbPixels), Me.ScaleY(Me.Height, vbTwips, vbPixels)

End Sub

Public Sub UpdateAndFormatFormula(Index As Integer, Optional blnForceUpdate As Boolean = False)
    Dim lngSaveLoc As Integer
    
    If mViewMode = vmdMultiView Or blnForceUpdate Then
        ' Only format if the multi-view frame is visible
        
        If rtfFormula(Index).Tag = FORMULA_HIGHLIGHTED Then
            rtfFormula(Index).Tag = FORMULA_CHANGED
        Else
            lngSaveLoc = rtfFormula(Index).SelStart
            
            If blnForceUpdate And mViewMode = vmdSingleView Then
                fraSingle.Visible = False
                fraMulti.Visible = True
            End If
            
            If mErrorParams.ErrorID <> 0 Then
                rtfFormula(Index).TextRTF = objMwtWin.TextToRTF(rtfFormula(Index).Text, False, True, True, mErrorParams.ErrorID)
            Else
                rtfFormula(Index).TextRTF = objMwtWin.TextToRTF(rtfFormula(Index).Text, False, True)
            End If
            rtfFormula(Index).SelStart = lngSaveLoc
            
            If blnForceUpdate And mViewMode = vmdSingleView Then
                fraSingle.Visible = True
                fraMulti.Visible = False
            End If
            
            If rtfFormula(Index).Tag = "" Then
                ' Only set as changed if it's blank
                ' Must preserve recalc and highlighted tags
                rtfFormula(Index).Tag = FORMULA_CHANGED
            End If
        End If
        If mViewMode = vmdMultiView Then
            rtfFormulaSingle.TextRTF = rtfFormula(Index).TextRTF
        End If
    End If

End Sub

Private Sub UpdateAndFormatFormulaSingle()
    Dim intSaveLoc As Integer
    Static blnUpdating As Boolean
    
    If blnUpdating Then Exit Sub
    blnUpdating = True
    
    If mViewMode = vmdSingleView Then
        ' Only format if the single-view frame is visible
        If rtfFormulaSingle.Tag = FORMULA_HIGHLIGHTED Then
            rtfFormulaSingle.Tag = FORMULA_CHANGED
            rtfFormula(mCurrentFormulaIndex).Tag = FORMULA_CHANGED
        Else
            intSaveLoc = rtfFormulaSingle.SelStart
            
            rtfFormulaSingle.TextRTF = objMwtWin.TextToRTF(rtfFormulaSingle.Text)
            rtfFormulaSingle.SelStart = intSaveLoc
            
            If rtfFormulaSingle.Tag = "" Then
                ' Only set as changed if it's blank
                ' Must preserve recalc and highlighted tags
                rtfFormulaSingle.Tag = FORMULA_CHANGED
                rtfFormula(mCurrentFormulaIndex).Tag = FORMULA_CHANGED
            End If
        End If
        rtfFormula(mCurrentFormulaIndex).TextRTF = rtfFormulaSingle.TextRTF
    End If
    
    blnUpdating = False

End Sub

Private Sub UpdateGrid(ThisGrid As MSFlexGrid, intFormulaIndex As Integer)

    Dim intIndex As Integer, intCurrentRow As Integer, intCurrentColumn As Integer
    Dim intColsNeeded As Integer
    Dim intRowSaved As Integer, intColumnSaved As Integer
    Dim lngWidthForTwoColumns As Long, lngWidthIncrement As Long
    Dim intCharLoc As Integer
    Dim lngDesiredValue As Long
    Dim strNewText As String
    
    ' First, set grid size and rows, then guess at number of columns
    
    With fraSingle
        lngDesiredValue = Me.ScaleHeight - fraSingle.Top - lblStatus.Height - 120
        If lngDesiredValue < 3000 Then lngDesiredValue = 3000
        .Height = lngDesiredValue
    End With
    
    With ThisGrid
        .Visible = False
        .Width = rtfFormulaSingle.Width + lblFormulaSingle.Width
        
        lngDesiredValue = fraSingle.Height - .Top - 360
        If lngDesiredValue < 0 Then lngDesiredValue = 0
        .Height = lngDesiredValue
        
        mPercentCompositionGridRowCount = .Height / .RowHeight(0) - 1
        If mPercentCompositionGridRowCount < 5 Then mPercentCompositionGridRowCount = 5
        .Rows = mPercentCompositionGridRowCount
        
        intRowSaved = .Row
        intColumnSaved = .Col
    
        If Not gBlnShowStdDevWithMass Or objMwtWin.StdDevMode = smShort Then
            lngWidthForTwoColumns = 6000
            lngWidthIncrement = 2000
        Else
            lngWidthForTwoColumns = 7000
            lngWidthIncrement = 2500
        End If
        
        If .Width < lngWidthForTwoColumns Then
            .Cols = 2
        Else
            .Cols = RoundToNearest(.Width / lngWidthIncrement, 1, False)
        End If
        
        For intIndex = 0 To .Cols - 1
            Select Case .Cols
            Case 2
                .ColWidth(intIndex) = (.Width - 120) / (.Cols) ' 30 twips is gridline width and/or height
            Case 3
                .ColWidth(intIndex) = (.Width - 40) / (.Cols) ' 30 twips is gridline width and/or height
            Case Else
                .ColWidth(intIndex) = (.Width - 40) / (.Cols) ' 30 twips is gridline width and/or height
            End Select
        Next intIndex
        
        .ScrollBars = flexScrollBarHorizontal
        
        ' Now, enlarge grid if necessary
        intColsNeeded = CIntSafe(mPctCompValues(intFormulaIndex).TotalElements / mPercentCompositionGridRowCount)
        If mPctCompValues(intFormulaIndex).TotalElements Mod mPercentCompositionGridRowCount > 0 Then
            intColsNeeded = intColsNeeded + 1
        End If
        If intColsNeeded > .Cols Then
            .Cols = intColsNeeded
            ' FrmMain.Width
            For intIndex = 1 To .Cols - 1
                .ColWidth(intIndex) = .ColWidth(0)
            Next intIndex
            .Height = (TextHeight("123456789gT") + 33) * (.Rows + 1) + 255      ' 255 is Scroll Bar Height
        Else
            .Height = (TextHeight("123456789gT") + 33) * (.Rows + 1)
        End If
        
        .LeftCol = 0
        .TopRow = 0
        
        ' Update mKeyPressAbortPctSolver if necessary
        If mKeyPressAbortPctSolver = 3 Then mKeyPressAbortPctSolver = 0
        If mKeyPressAbortPctSolver = 2 Then mKeyPressAbortPctSolver = 3
        
    End With
    
    ' Finally, copy the data in .GridCell().Text into the grid
    intCurrentRow = 0
    intCurrentColumn = 0
    For intIndex = 1 To mPctCompValues(intFormulaIndex).TotalElements
        With mPctCompValues(intFormulaIndex).GridCell(intIndex)
            If .Locked And PctSolverParams.SolverOn And mKeyPressAbortPctSolver = 0 Then
                intCharLoc = InStr(.Text, ":")
                If intCharLoc > 0 Then
                     strNewText = SpacePad(Left(.Text, intCharLoc), 4) & CStr(.Goal)
                Else
                    ' Colon not found; this is unexpected
                    Debug.Assert False
                    strNewText = SpacePad(Left(.Text, 3), 4) & CStr(.Goal)
                End If
            Else
                strNewText = .Text
            End If
            ThisGrid.TextMatrix(intCurrentRow, intCurrentColumn) = strNewText
        End With
        
        intCurrentRow = intCurrentRow + 1
        If intCurrentRow = ThisGrid.Rows Then
            intCurrentRow = 0
            intCurrentColumn = intCurrentColumn + 1
        End If
    Next intIndex
    
    With grdPC
        ' Erase remaining cells
        For intIndex = mPctCompValues(intFormulaIndex).TotalElements + 1 To .Rows * .Cols
            .TextMatrix(intCurrentRow, intCurrentColumn) = ""
            intCurrentRow = intCurrentRow + 1
            If intCurrentRow = ThisGrid.Rows Then
                If intCurrentColumn <> .Cols - 1 Then
                    intCurrentRow = 0
                    intCurrentColumn = intCurrentColumn + 1
                Else
                    ' Last cell reached, for loop will exit on next intIndex
                    ' Just keep cell in same spot
                End If
            End If
        Next intIndex
        
        ' Show the grid again
        .Visible = True
        
        ' Re-position cursor
        If intRowSaved >= .Rows Then intRowSaved = 0
        If intColumnSaved >= .Cols Then intColumnSaved = 0
        
        .Row = intRowSaved
        .Col = intColumnSaved
        
    End With
    
End Sub

Private Sub cmdCalculate_Click()
    InitiateCalculate
End Sub

Private Sub cmdNewFormula_Click()
    AddNewFormulaWrapper
End Sub

Private Sub Form_Click()
    If mKeyPressAbortPctSolver = 1 Then
        ' Key pressed, so end calculations
        mKeyPressAbortPctSolver = 2
    End If
End Sub

Private Sub Form_Deactivate()
    If Not mIgnoreNextGotFocus Then
        Calculate True, True, True, 0, False, False, False, 1, False
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    ' Since KeyPreview is true for frmMain, the Form_Keydown event gets called before the
    '   control_keydown event
    ' Useful for catching F1 and Ctrl+A keys
    
    FormWideKeyDownHandler KeyCode, Shift
'    lblStatus.Caption = lblStatus.Caption & "   " & Str(keycode)  ' Useful for finding a key's Keycode

End Sub

Private Sub Form_Load()
    
    Dim eResponse As VbMsgBoxResult
    Dim blnSuccess As Boolean
    Dim blnShowDebugPrompts As Boolean
    Dim lngErrorID As Long, blnProceedWithLoad As Boolean
    Dim intInstances As Integer
    
    On Error GoTo ErrorHandler
    
    ' Uncomment the following to Debug
    If InStr(UCase(Command()), "/DEBUG") Then
        eResponse = MsgBox("Track load progress?", vbYesNo + vbDefaultButton2, "Track Progress")
        If eResponse = vbYes Then blnShowDebugPrompts = True
    End If
    
    If blnShowDebugPrompts Then MsgBox "Initialize globals"
    InitializeFormWideVariables
    InitializeGlobalVariables blnShowDebugPrompts
    
    ' Set frmMain.lblHiddenFormStatus to "Loading"
    ' Used by Functions to prevent certain tasks from occuring while program is loading
    frmMain.lblHiddenFormStatus = "Loading"
    
    If blnShowDebugPrompts Then MsgBox "Show intro form"
    frmIntro.Show
    frmIntro.MousePointer = vbHourglass
    DoEvents
    
    If blnShowDebugPrompts Then MsgBox "Set default options"
    ' Load default options and set default menu checkmarks
    SetDefaultOptions
    
    If blnShowDebugPrompts Then MsgBox "Set default values and formulas"
    SetDefaultValuesAndFormulas False, blnShowDebugPrompts
    
    If blnShowDebugPrompts Then MsgBox "Load string constants"
    ' Load String Constants into MSFlexGrids on frmStrings
    MemoryLoadAllStringConstants
    
    If blnShowDebugPrompts Then MsgBox "Size and center App"
    ' Put app in center of screen horizontally, and upper fourth vertically
    SizeAndCenterWindow Me, cWindowUpperThird, 6300, MAIN_HEIGHT
    
    If blnShowDebugPrompts Then MsgBox "Position controls"
    ' Position frmMain controls
    PositionFormControls
    
    If blnShowDebugPrompts Then MsgBox "Pre-Parse command line"
    ' Preliminarily parse the command line
    ParseCommandLine True
    
    If blnShowDebugPrompts Then MsgBox "Load forms FinderModeWarn and IonMatchOptions"
    ' Load various forms that need to be loaded (but not shown) so Language-specific
    '  captions can be set
    Load frmFinderModeWarn
    Load frmIonMatchOptions
    Load frmDtaTxtFileBrowser
    
    frmFinderModeWarn.Visible = False
    frmIonMatchOptions.Visible = False
    frmDtaTxtFileBrowser.Visible = False
    DoEvents
    
    If blnShowDebugPrompts Then MsgBox "Load program options from disk"
    ' Load Default Program Options (from disk, or memory if not found)
    LoadDefaultOptions False, blnShowDebugPrompts
    
    ' If gCurrentLanguage is not English, then attempt to load correct
    '  language file
    If LCase(gCurrentLanguageFileName) <> LCase(DEFAULT_LANGUAGE_FILENAME) Then
        If blnShowDebugPrompts Then MsgBox "Load language file from " & gCurrentLanguageFileName
        blnSuccess = LoadLanguageSettings(gCurrentLanguageFileName, gCurrentLanguage)
        If Not blnSuccess Then
            ' If alternate langauge file could not be loaded, then revert to English
            gCurrentLanguageFileName = DEFAULT_LANGUAGE_FILENAME
            gCurrentLanguage = "English"
            SaveSingleDefaultOption "Language", gCurrentLanguage
            SaveSingleDefaultOption "LanguageFile", gCurrentLanguageFileName
        End If
    Else
        ' Set blnSuccess to False to assure Menu items get loaded from memory
        blnSuccess = False
    End If
    
    If blnShowDebugPrompts Then MsgBox "Reset menu captions"
    If Not blnSuccess Then
        ResetMenuCaptions True
    End If
    
    ' Now that Language Information has been loaded, check if other copies
    '  of the Molecular Weight Calculator are running.  If they are, prompt user
    '  about continuing
    
    ' Assume OK to proceed
    blnProceedWithLoad = True
    intInstances = CountInstancesOfApp(Me)
    If intInstances > 1 Then
        eResponse = YesNoBox(LookupLanguageCaption(5350, "The Molecular Weight Calculator is already running.  Are you sure you want to start another copy?"), LookupLanguageCaption(5355, "Already Running"))
        If eResponse <> vbYes Then blnProceedWithLoad = False
    End If
    
    If Not blnProceedWithLoad Then
        ' Exit program, but do not save values or default options
        gNonSaveExitApp = True
    Else
        If blnShowDebugPrompts Then MsgBox "Parse command line"
        ' Modify defaults based on command line
        gCommandLineInstructionsDisplayed = ParseCommandLine(False)
            
        If gCommandLineInstructionsDisplayed Then
            ' Also set gNonSaveExitApp to true so frmMain won't auto unload, allowing user
            ' To read message on frmIntro
            gNonSaveExitApp = True
        Else
            If blnShowDebugPrompts Then MsgBox "Set weight mode"
            SwitchWeightMode emAverageMass
            
            ' Load elements from disk
            LoadElements 0, True

            ' Now that elements are loaded, need to populate grid on Edit Elements form with elements
            frmEditElem.PopulateGrid
            
            If blnShowDebugPrompts Then MsgBox "Load abbreviations from disk"
            ' Load all abbreviations (from disk, or memory if not found)
            LoadAbbreviations False
        
            ' Show the main form even if gBlnLoadStatusOK = False
            frmMain.Show
            
            If mBlnBatchProcess Then
                BatchProcessTextFile mBatchProcessFilename, mBatchProcessOutfile, mBatchProcessOverwrite
                'lblStatus.Caption = "Completed processing " & mBatchProcessFilename
                End
            End If
                
            If blnShowDebugPrompts Then MsgBox "Load default values from disk"
            ' Load Default Values (from disk or memory)
            LoadValuesAndFormulas False
            
            mCurrentFormulaIndex = 0
            lblFormulaSingle.Caption = lblFormula(mCurrentFormulaIndex).Caption
            
            ' Set frmMain.lblHiddenFormStatus to "Done Loading"
            frmMain.lblHiddenFormStatus = "Done Loading"
        
            ' This info is used by the Form_Resize event
            frmProgramPreferences.lblHiddenDefaultsLoadedStatus = "DefaultsLoaded"
            ResizeFormMain True
            
            If blnShowDebugPrompts Then MsgBox "ReCalculate"
            
            ' ReCalculate
            lngErrorID = Calculate(True)
                        
            ' Fill in Status message
            LabelStatus lngErrorID
            
            ' Reset mouse pointer
            frmIntro.MousePointer = vbDefault
            
            ' Hide Intro box if no problems
            If gBlnLoadStatusOK Then
                If blnShowDebugPrompts Then MsgBox "Unload intro form"
                Unload frmIntro
            
                ' Set focus to rtfFormula(0) only if gBlnLoadStatusOK = True
                SetFocusToFormulaByIndex
            End If
            
            ' Display default form if request by user (and not Main form)
            ShowDefaultFormAtLoad
            
            If blnShowDebugPrompts Then MsgBox "Done Loading"
        End If
    End If
    
ProgStart:
    Exit Sub

ErrorHandler:
    Close
    GeneralErrorHandler "frmMain|Form_Load", Err.Number, Err.Description
    Resume ProgStart

End Sub

Private Sub Form_Resize()
    ResizeFormMain
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim intIndex As Integer
    
    ' Set the following two variables to 2 to stop the Formula finder and Pct Solver (if running)
    gKeyPressAbortFormulaFinder = 2
    mKeyPressAbortPctSolver = 2

    If cChkBox(frmProgramPreferences.chkAutosaveValues) And Not gNonSaveExitApp Then
        SaveDefaultOptions
        SaveValuesAndFormulas
        If objMwtWin.GetElementMode <> gElementWeightTypeInFile Then
            ' User quick switched the elementweight type, but it hasn't been saved
            ' Save it Now
            SaveElements
        End If
    End If
    
    ' frmStrings and frmSetValue Refuse to unload using the code in UnloadAllForms
    ' Manually unload if loaded
    If IsLoaded("frmSetValue") Then Unload frmSetValue
    If IsLoaded("frmStrings") Then Unload frmStrings

    ' Unload the objCompounds()
    For intIndex = 0 To MAX_FORMULAS
        Set objCompounds(intIndex) = Nothing
    Next intIndex
    
    ' Unload the objMwtWin class
    Set objMwtWin = Nothing
    
    UnloadAllForms Me.Name
    
    End
    
End Sub

Private Sub fraSingle_Click()
    If mKeyPressAbortPctSolver = 1 Then
        ' Key pressed, so end calculations
        mKeyPressAbortPctSolver = 2
    End If

End Sub

Private Sub grdPC_GotFocus()
    mUsingPCGrid = True
End Sub

Private Sub grdPC_KeyDown(KeyCode As Integer, Shift As Integer)

    If mKeyPressAbortPctSolver = 1 Then
        ' Key pressed, so end calculations
        mKeyPressAbortPctSolver = 2
    End If
    
    Select Case KeyCode
    Case vbKeyReturn
        grdPC_Click
    Case vbKeyPageUp, vbKeyUp
        ' Page Up and up arrow
        If grdPC.Row = 0 Then
            rtfFormulaSingle.SetFocus
            mUsingPCGrid = False
            LabelStatus
        End If
    Case vbKeyPageDown, vbKeyDown
        ' Page Down and down arrow
        If grdPC.Row = 10 Then
            rtfFormulaSingle.SetFocus
            mUsingPCGrid = False
            LabelStatus
        End If
    End Select

End Sub

Private Sub grdPC_LostFocus()
    mUsingPCGrid = False
End Sub

Private Sub lblFormulaSingle_Click()
    If mKeyPressAbortPctSolver = 1 Then
        ' Key pressed, so end calculations
        mKeyPressAbortPctSolver = 2
    End If
End Sub

Private Sub txtMWT_Change(Index As Integer)
    txtMWTSingle.Text = txtMWT(Index).Text
End Sub

Private Sub txtMWTSingle_Click()
    If mKeyPressAbortPctSolver = 1 Then
        ' Key pressed, so end calculations
        mKeyPressAbortPctSolver = 2
    End If
End Sub

Private Sub lblStatus_Click()
    If mKeyPressAbortPctSolver = 1 Then
        ' Key pressed, so end calculations
        mKeyPressAbortPctSolver = 2
    End If
End Sub

Private Sub lblStatus_DblClick()
    ZoomLine lblStatus.Caption
End Sub

Private Sub lblValueForX_Click()
    If mKeyPressAbortPctSolver = 1 Then
        ' Key pressed, so end calculations
        mKeyPressAbortPctSolver = 2
    End If
End Sub

Private Sub mnuAbout_Click()
    frmAboutBox.Show vbModal
End Sub

Private Sub mnuAminoAcidNotationConverter_Click()
    ShowAminoAcidNotationModule
End Sub

Private Sub mnuChooseLanguage_Click()
    frmChooseLanguage.Show
End Sub

Private Sub mnuIsotopicDistribution_Click()
    ShowIsoDistributionModeller
End Sub

Private Sub mnuPeptideSequenceFragmentation_Click()
    ShowPeptideSequenceModeller
End Sub

Private Sub mnuDisplayIsotopicDistribution_Click()
    DisplayIsotopicDistribution mCurrentFormulaIndex
End Sub

Private Sub mnuCalculateFile_Click()
    ' Call the BatchProcessing subroutine
    BatchProcessTextFile
End Sub

Private Sub mnuCalculator_Click()
    frmCalculator.Show
End Sub

Private Sub mnuCapillaryFlow_Click()
    ShowCapillaryFlowModule
End Sub

Private Sub mnuChangeFont_Click()

    frmChangeFont.Show

End Sub

Private Sub mnuCopy_Click()
    CopyRoutine Me, True
End Sub

Private Sub mnuCopyCurrent_Click()
    DuplicateCurrentFormula
End Sub

Private Sub mnuCopyMWT_Click()
    CopyCurrentMWT False
End Sub

Private Sub mnuCopyPC_Click()
    CopyPC
End Sub

Private Sub mnuCopyRTF_Click()
    CopyRTF
End Sub

Private Sub mnuCut_Click()
    CutRoutine Me, True
End Sub

Private Sub mnuDelete_Click()
    ' Delete selected text.
    frmMain.ActiveControl.SelText = ""
End Sub

Private Sub mnuEdit_Click()
    EditMenuCheck
End Sub

Private Sub mnuEditAbbrev_Click()
    EditAbbreviations
End Sub

Private Sub mnuEditElements_Click()
    EditElements
End Sub

Private Sub mnuEmpirical_Click()
    ConvertToEmpirical
End Sub

Private Sub mnuEraseAll_Click()
    EraseAllFormulas True
End Sub

Private Sub mnuEraseCurrent_Click()
    EraseCurrentFormula
End Sub

Private Sub mnuExit_Click()
    ExitProgram
End Sub

Private Sub mnuExpandAbbrev_Click()
    ExpandAbbreviations
End Sub

Private Sub mnuFinder_Click()
    ShowFormulaFinder
End Sub

Private Sub mnuMMConvert_Click()
    ShowMoleMassConverter
End Sub

Private Sub mnuOverview_Click()
    On Error GoTo ErrHandler

    'hWnd is a Long defined elsewhere to be the window handle
    'that will be the parent to the help window.
    Dim hwndHelp As Long
    
    'The return value is the window handle of the created help window.
    hwndHelp = HtmlHelp(hwnd, App.HelpFile, HH_DISPLAY_TOPIC, 0)
    
    Exit Sub
    
ErrHandler:
    ' User pressed cancel button
    Exit Sub

End Sub

Private Sub mnuPaste_Click()
    PasteRoutine Me, True
End Sub

Private Sub mnuPercentType_Click(Index As Integer)
    If Index = psmPercentSolverOn Then
        TogglePercentComposition psmPercentSolverOn
    Else
        TogglePercentComposition psmPercentSolverOff
    End If
End Sub

Private Sub mnuPrint_Click()
    PrintResults
End Sub

Private Sub ToggleToolTips()
    frmProgramPreferences.SwapCheck frmProgramPreferences.chkShowToolTips
End Sub

Private Sub mnuProgramOptions_Click()
    ShowProgramOptions
End Sub

Private Sub mnuRestoreValues_Click()
    RestoreDefaultValues
End Sub

Private Sub mnuRightClickCopy_Click()
    CopyRoutine Me, True
End Sub

Private Sub mnuRightClickCut_Click()
    CutRoutine Me, True
End Sub

Private Sub mnuRightClickDelete_Click()
    frmMain.ActiveControl.SelText = ""
End Sub

Private Sub mnuRightClickPaste_Click()
    PasteRoutine Me, True
End Sub

Private Sub mnuRightClickSelectAll_Click()
    frmMain.ActiveControl.SelStart = 0
    frmMain.ActiveControl.SelLength = Len(frmMain.ActiveControl.Text)
End Sub

Private Sub mnuRightClickUndo_Click()
    If TypeOf frmMain.ActiveControl Is RichTextBox Or _
       TypeOf frmMain.ActiveControl Is VB.TextBox Then
        frmMain.ActiveControl.Text = GetMostRecentTextBoxValue()
    End If
End Sub

Private Sub mnuSaveValues_Click()
    SaveValuesAndFormulas
End Sub

Private Sub mnuShowTips_Click()
    ToggleToolTips
End Sub

Private Sub mnuStayOnTop_Click()
    ToggleStayOnTop
End Sub

Private Sub mnuViewType_Click(Index As Integer)
    If Index = vmdSingleView Then
        SetViewMode vmdSingleView
    Else
        SetViewMode vmdMultiView
    End If
    
End Sub

Private Sub grdPC_Click()
    PercentCompGridClickHandler
End Sub

Private Sub optElementMode_Click(Index As Integer)
    ' Elementweightmode = 1 means average weights, 2 means isotopic, and 3 means integer
    Dim eNewWeightMode As emElementModeConstants
    
    eNewWeightMode = Index + 1
        
    If eNewWeightMode <> objMwtWin.GetElementMode Then
        SwitchWeightMode eNewWeightMode
    End If

End Sub

Private Sub rtfFormula_Change(Index As Integer)
    UpdateAndFormatFormula Index
End Sub

Private Sub rtfFormula_GotFocus(Index As Integer)
    lblFormulaSingle.Caption = ConstructFormulaLabel(Index)
    mCurrentFormulaIndex = Index
    SetMostRecentTextBoxValue rtfFormula(Index).Text
End Sub

Private Sub rtfFormula_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    ' Look for arrow keys or Enter key
    FormulaKeyDownHandler KeyCode, Shift
End Sub

Private Sub rtfFormula_KeyPress(Index As Integer, KeyAscii As Integer)
    ' Make sure the key is valid
    FormulaKeyPressHandler rtfFormula(Index), KeyAscii
End Sub

Private Sub rtfFormula_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        ShowPopupRightClickMenu Me, rtfFormula(Index), True
    End If
End Sub

Private Sub rtfFormulaSingle_Change()
    UpdateAndFormatFormulaSingle
End Sub

Private Sub rtfFormulaSingle_Click()
    If mKeyPressAbortPctSolver = 1 Then
        ' Key pressed, so end calculations
        mKeyPressAbortPctSolver = 2
    End If
End Sub

Private Sub rtfFormulaSingle_GotFocus()
    If Not mIgnoreNextGotFocus Then
        If Not gBlnErrorPresent Then
            UpdateGrid grdPC, mCurrentFormulaIndex
        End If
        
        ' The .SetFocus event can cause an error if a form is shown modally
        ' Must use Resume Next error handling to avoid an error
        On Error Resume Next
        rtfFormulaSingle.SetFocus
        On Error GoTo 0
        
        SetMostRecentTextBoxValue rtfFormulaSingle.Text
    Else
        mIgnoreNextGotFocus = False
    End If
End Sub

Private Sub rtfFormulaSingle_KeyDown(KeyCode As Integer, Shift As Integer)
    FormulaKeyDownHandler KeyCode, Shift
End Sub

Private Sub rtfFormulaSingle_KeyPress(KeyAscii As Integer)
    FormulaKeyPressHandler rtfFormulaSingle, KeyAscii
End Sub
