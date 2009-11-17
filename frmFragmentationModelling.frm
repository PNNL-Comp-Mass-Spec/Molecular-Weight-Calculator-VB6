VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form frmFragmentationModelling 
   Caption         =   "Peptide Sequence Fragmentation Modelling"
   ClientHeight    =   8640
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   10020
   HelpContextID   =   3080
   Icon            =   "frmFragmentationModelling.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8640
   ScaleWidth      =   10020
   StartUpPosition =   3  'Windows Default
   Tag             =   "12000"
   Begin VB.CommandButton cmdMatchIons 
      Caption         =   "&Match Ions"
      Height          =   345
      Left            =   7200
      TabIndex        =   2
      Tag             =   "12020"
      Top             =   240
      Width           =   1515
   End
   Begin VB.Frame fraMassInfo 
      Caption         =   "Mass Information"
      Height          =   935
      Left            =   2640
      TabIndex        =   4
      Tag             =   "12410"
      Top             =   720
      Width           =   7275
      Begin VB.OptionButton optElementMode 
         Caption         =   "Average"
         Height          =   255
         HelpContextID   =   4053
         Index           =   0
         Left            =   5640
         TabIndex        =   55
         Tag             =   "12425"
         Top             =   360
         Value           =   -1  'True
         Width           =   1395
      End
      Begin VB.OptionButton optElementMode 
         Caption         =   "Isotopic"
         Height          =   255
         HelpContextID   =   4053
         Index           =   1
         Left            =   5640
         TabIndex        =   54
         Tag             =   "12430"
         Top             =   600
         Width           =   1395
      End
      Begin VB.TextBox txtMWT 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   53
         Text            =   "MW="
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox txtMHAlt 
         Height          =   285
         Left            =   3720
         TabIndex        =   9
         Text            =   "500"
         Top             =   560
         Width           =   1335
      End
      Begin VB.ComboBox cboMHAlternate 
         Height          =   315
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   540
         Width           =   1215
      End
      Begin VB.TextBox txtMH 
         Height          =   285
         Left            =   3720
         TabIndex        =   6
         Text            =   "500"
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblElementMode 
         Caption         =   "Element Mode"
         Height          =   225
         Left            =   5640
         TabIndex        =   56
         Tag             =   "12420"
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label lblMHAltDa 
         Caption         =   "Da"
         Height          =   255
         Left            =   5160
         TabIndex        =   10
         Tag             =   "12350"
         Top             =   585
         Width           =   375
      End
      Begin VB.Label lblMHDa 
         Caption         =   "Da"
         Height          =   255
         Left            =   5160
         TabIndex        =   7
         Tag             =   "12350"
         Top             =   255
         Width           =   375
      End
      Begin VB.Label lblMH 
         Caption         =   "[M+H]1+"
         Height          =   255
         Left            =   2400
         TabIndex        =   5
         Top             =   260
         Width           =   855
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdIonList 
      Height          =   6375
      Left            =   6960
      TabIndex        =   46
      Top             =   1800
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   11245
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
   End
   Begin VB.Frame fraIonStats 
      Caption         =   "Ion Statistics"
      Height          =   1935
      Left            =   3960
      TabIndex        =   47
      Tag             =   "12360"
      Top             =   6480
      Width           =   3015
      Begin VB.Label lblScore 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   1640
         Width           =   2655
      End
      Begin VB.Label lblMatchCount 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   1400
         Width           =   2655
      End
      Begin VB.Label lblBinAndToleranceCounts 
         BorderStyle     =   1  'Fixed Single
         Height          =   615
         Left            =   120
         TabIndex        =   50
         Top             =   440
         Width           =   2655
      End
      Begin VB.Label lblPrecursorStatus 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   120
         TabIndex        =   49
         Top             =   1040
         Width           =   2655
      End
      Begin VB.Label lblIonLoadedCount 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   200
         Width           =   2655
      End
   End
   Begin VB.Frame fraIonMatching 
      Caption         =   "Ion Match Options"
      Height          =   1935
      Left            =   60
      TabIndex        =   31
      Tag             =   "12300"
      Top             =   6480
      Width           =   3735
      Begin VB.TextBox txtAlignment 
         Height          =   285
         Left            =   2280
         TabIndex        =   43
         Tag             =   "12060"
         Text            =   "0"
         Top             =   1560
         Width           =   735
      End
      Begin VB.CheckBox chkRemovePrecursorIon 
         Caption         =   "&Remove Precursor Ion"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Tag             =   "12310"
         Top             =   240
         Value           =   1  'Checked
         Width           =   3375
      End
      Begin VB.TextBox txtPrecursorMassWindow 
         Height          =   285
         Left            =   2280
         TabIndex        =   37
         Text            =   "2"
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txtIonMatchingWindow 
         Height          =   285
         Left            =   2280
         TabIndex        =   40
         Text            =   ".5"
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtPrecursorIonMass 
         Height          =   285
         Left            =   2280
         TabIndex        =   34
         Text            =   "500"
         Top             =   500
         Width           =   735
      End
      Begin VB.Label lblAlignment 
         Caption         =   "Alignment &Offset"
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Tag             =   "12355"
         Top             =   1580
         Width           =   2175
      End
      Begin VB.Label lblDaltons 
         Caption         =   "Da"
         Height          =   255
         Index           =   3
         Left            =   3120
         TabIndex        =   44
         Tag             =   "12350"
         Top             =   1580
         Width           =   495
      End
      Begin VB.Label lblDaltons 
         Caption         =   "Da"
         Height          =   255
         Index           =   2
         Left            =   3120
         TabIndex        =   41
         Tag             =   "12350"
         Top             =   1220
         Width           =   495
      End
      Begin VB.Label lblDaltons 
         Caption         =   "Da"
         Height          =   255
         Index           =   1
         Left            =   3120
         TabIndex        =   38
         Tag             =   "12350"
         Top             =   860
         Width           =   495
      End
      Begin VB.Label lblPrecursorMassWindow 
         Caption         =   "Mass Window"
         Height          =   255
         Left            =   480
         TabIndex        =   36
         Tag             =   "12330"
         Top             =   860
         Width           =   1815
      End
      Begin VB.Label lblDaltons 
         Caption         =   "Da"
         Height          =   255
         Index           =   0
         Left            =   3120
         TabIndex        =   35
         Tag             =   "12350"
         Top             =   520
         Width           =   495
      End
      Begin VB.Label lblIonMatchingWindow 
         Caption         =   "&Ion Matching Window"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Tag             =   "12340"
         Top             =   1220
         Width           =   2175
      End
      Begin VB.Label lblPrecursorIonMass 
         Caption         =   "Ion Mass"
         Height          =   255
         Left            =   480
         TabIndex        =   33
         Tag             =   "12320"
         Top             =   520
         Width           =   1695
      End
   End
   Begin VB.Frame fraCharge 
      Caption         =   "Charge Options"
      Height          =   1095
      Left            =   60
      TabIndex        =   25
      Tag             =   "12260"
      Top             =   5280
      Width           =   2415
      Begin VB.ComboBox cboTripleCharge 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   30
         ToolTipText     =   "The 2+ m/z value will be computed for ions above this m/z"
         Top             =   720
         Width           =   1215
      End
      Begin VB.ComboBox cboDoubleCharge 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Tag             =   "12285"
         ToolTipText     =   "The 2+ m/z value will be computed for ions above this m/z"
         Top             =   360
         Width           =   1215
      End
      Begin VB.CheckBox chkTripleCharge 
         Caption         =   "&3+ ions"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Tag             =   "12272"
         Top             =   720
         Width           =   950
      End
      Begin VB.CheckBox chkDoubleCharge 
         Caption         =   "&2+ ions"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Tag             =   "12270"
         Top             =   360
         Width           =   950
      End
      Begin VB.Label lblChargeThreshold 
         Caption         =   "Threshold"
         Height          =   255
         Left            =   1200
         TabIndex        =   26
         Tag             =   "12280"
         Top             =   120
         Width           =   1100
      End
   End
   Begin VB.Frame fraTerminii 
      Caption         =   "N and C Terminus"
      Height          =   1255
      Left            =   60
      TabIndex        =   11
      Tag             =   "12150"
      Top             =   1680
      Width           =   2415
      Begin VB.ComboBox cboCTerminus 
         Height          =   315
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Tag             =   "12190"
         Top             =   800
         Width           =   1575
      End
      Begin VB.ComboBox cboNTerminus 
         Height          =   315
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Tag             =   "12180"
         Top             =   320
         Width           =   1575
      End
      Begin VB.Label lblCTerminus 
         Caption         =   "&C"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Tag             =   "12170"
         Top             =   840
         Width           =   300
      End
      Begin VB.Label lblNTerminus 
         Caption         =   "&N"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Tag             =   "12160"
         Top             =   360
         Width           =   300
      End
   End
   Begin VB.Frame fraNeutralLosses 
      Caption         =   "Neutral Losses"
      Height          =   1005
      Left            =   60
      TabIndex        =   20
      Tag             =   "12230"
      Top             =   4200
      Width           =   2415
      Begin VB.CheckBox chkPhosphateLoss 
         Caption         =   "Loss of PO4"
         Height          =   255
         Left            =   840
         TabIndex        =   24
         Tag             =   "12255"
         Top             =   680
         Width           =   1500
      End
      Begin VB.ListBox lstIonsToModify 
         Height          =   645
         ItemData        =   "frmFragmentationModelling.frx":08CA
         Left            =   120
         List            =   "frmFragmentationModelling.frx":08DD
         MultiSelect     =   1  'Simple
         TabIndex        =   21
         Tag             =   "12235"
         ToolTipText     =   "Choose ions to which losses will be applied"
         Top             =   240
         Width           =   550
      End
      Begin VB.CheckBox chkAmmoniaLoss 
         Caption         =   "Loss of NH3"
         Height          =   255
         Left            =   840
         TabIndex        =   23
         Tag             =   "12250"
         Top             =   440
         Width           =   1500
      End
      Begin VB.CheckBox chkWaterLoss 
         Caption         =   "Loss of H2O"
         Height          =   255
         Left            =   840
         TabIndex        =   22
         Tag             =   "12240"
         Top             =   200
         Value           =   1  'Checked
         Width           =   1500
      End
   End
   Begin VB.Frame fraIonTypes 
      Caption         =   "Ion Types"
      Height          =   1095
      Left            =   60
      TabIndex        =   16
      Tag             =   "12200"
      Top             =   3000
      Width           =   2415
      Begin VB.CheckBox chkIonType 
         Caption         =   "C Ions"
         Height          =   255
         Index           =   3
         Left            =   1200
         TabIndex        =   58
         Tag             =   "12215"
         Top             =   480
         Width           =   1000
      End
      Begin VB.CheckBox chkIonType 
         Caption         =   "Z Ions"
         Height          =   255
         Index           =   4
         Left            =   1200
         TabIndex        =   57
         Tag             =   "12220"
         Top             =   720
         Width           =   1000
      End
      Begin VB.CheckBox chkIonType 
         Caption         =   "&Y Ions"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   19
         Tag             =   "12220"
         Top             =   720
         Value           =   1  'Checked
         Width           =   1000
      End
      Begin VB.CheckBox chkIonType 
         Caption         =   "&B Ions"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   18
         Tag             =   "12215"
         Top             =   480
         Value           =   1  'Checked
         Width           =   1000
      End
      Begin VB.CheckBox chkIonType 
         Caption         =   "&A Ions"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Tag             =   "12210"
         Top             =   240
         Width           =   1000
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdFragMasses 
      Height          =   4455
      Left            =   2520
      TabIndex        =   45
      Top             =   1800
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   7858
      _Version        =   393216
      FixedCols       =   0
   End
   Begin VB.ComboBox cboNotation 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Tag             =   "12010"
      ToolTipText     =   "Amino acid sequence notation type"
      Top             =   840
      Width           =   2295
   End
   Begin VB.TextBox txtSequence 
      Height          =   555
      Left            =   1440
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Tag             =   "12050"
      ToolTipText     =   "Enter amino acid sequence here"
      Top             =   120
      Width           =   5535
   End
   Begin VB.Label lblSequence 
      Caption         =   "&Sequence:"
      Height          =   255
      Left            =   60
      TabIndex        =   0
      Tag             =   "12100"
      Top             =   180
      Width           =   1335
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuLoadSequenceInfo 
         Caption         =   "&Load Sequence Info"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuSaveSequenceInfo 
         Caption         =   "&Save Sequence Info"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLoadIonList 
         Caption         =   "Load List of &Ions or .Dta file to Match"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuCopyPredictedIons 
         Caption         =   "&Copy Predicted Ions"
      End
      Begin VB.Menu mnuCopyPredictedIonsAsRTF 
         Caption         =   "Copy Predicted Ions as &RTF"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuCopyPredictedIonsAsHtml 
         Caption         =   "Copy Predicted Ions as &Html"
      End
      Begin VB.Menu mnuEditSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCopySequenceMW 
         Caption         =   "Copy Sequence Molecular &Weight"
         Shortcut        =   ^W
      End
      Begin VB.Menu mnuEditSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPasteIonList 
         Caption         =   "&Paste List of Ions to Match"
      End
      Begin VB.Menu mnuClearMatchIonList 
         Caption         =   "Clear Match Ion &List"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewMatchIonList 
         Caption         =   "List of &Ions to Match"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuViewSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowMassSpectrum 
         Caption         =   "&Mass Spectrum"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuUpdateSpectrum 
         Caption         =   "&Update Spectrum on Change"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuViewDtaTxtSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewDtaTxtBrowser 
         Caption         =   "&Dta.Txt File Browser"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuIonMatchListOptions 
         Caption         =   "Ion Match List &Options"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuEditModificationSymbols 
         Caption         =   "Edit Residue &Modification Symbols"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuToolsSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAutoAlign 
         Caption         =   "&Automatically Align Ions to Match"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuFragmentationModellingHelp 
         Caption         =   "&Fragmentation Modelling"
      End
   End
   Begin VB.Menu mnuIonMatchListRightClick 
      Caption         =   "IonMatchListRightClickMenu"
      Begin VB.Menu mnuIonMatchListRightClickCopy 
         Caption         =   "&Copy Selected Ions"
      End
      Begin VB.Menu mniIonMatchListRightClickPaste 
         Caption         =   "&Paste List of Ions to Match"
      End
      Begin VB.Menu mnuIonMatchListRightClickDeleteAll 
         Caption         =   "&Delete All (Clear list)"
      End
      Begin VB.Menu mnuIonMatchListRightClickSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIonMatchListRightClickSelectAll 
         Caption         =   "Select &All"
      End
   End
End
Attribute VB_Name = "frmFragmentationModelling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const TOTAL_POSSIBLE_ION_TYPES = 5
Private Const MAX_CHARGE_STATE = 3
Private Const MAX_MODIFICATIONS = 6             ' Maximum number of modifications for a single residue
Private Const SHOULDER_ION_PREFIX = "Shoulder-"

Private Const ANNOTATION_FONT_SIZE = 10
Private Const MAX_SERIES_COUNT = 32

' The following values are now given by the itIonTypeConstants enumerated type
'Const itAIon = 0
'Const itBIon = 1
'Const itYIon = 2
'Const itCIon = 3
'Const itZIon = 4

Private Const MAX_BIN_COUNT = 50000
Private Const ION_SEPARATION_TOLERANCE_TO_AUTO_ENABLE_BINNING = 0.2

Const ION_LIST_COL_COUNT = 3
Private Enum ilgIonListGridConstants
    ilgMass = 0
    ilgIntensity = 1
    ilgSymbolMatch = 2
End Enum

Const ION_LIST_ARRAY_MAXINDEX_2D = 2
Private Enum ilaIonListArrayConstants
    ilaMass = 0
    ilaIntensity = 1
    ilaNormalizedIntensity = 2
End Enum

Private Type udtGridLocationType
    Row As Long
    Col As Long
End Type

' The udtResidueMatchedType type is used to record if the given ion type of the
'  given charge was matched in sub MatchIons
' For example, if the b ion for residue 2 is seen singly charged, then
'  .IonType(itBIon, 1) = True
Private Type udtResidueMatchedType
    IonHit(TOTAL_POSSIBLE_ION_TYPES, MAX_CHARGE_STATE) As Boolean
End Type

Private Type udtIonMatchListDetailsType
    Caption As String
    OriginalMaximumIntensity As Double
End Type

Private dblImmoniumMassDifference As Double ' CO minus H = 26.9871

Private dblHistidineFW As Double            ' 110
Private dblPhenylalanineFW As Double        ' 120
Private dblTyrosineFW As Double             ' 136

Private mFormActivatedByUser As Boolean      ' Set true the first time the user activates the form; always true from then on
                                             ' Used to prevent plotting the mass spectrum until after the user has activated the fragmentation match form at least once
                                             
Private mFragMatchSettingsChanged As Boolean

' FragSpectrumDetails() contains detailed information on the fragmentation spectrum data, sorted by mass
' FragSpectrumGridLocs() contains the row and column indices in grdFragMasses that the values in FragSpectrumDetails() are displayed at
' Both contain an equal number of data points, given by FragSpectrumIonCount
Private FragSpectrumIonCount As Long
Private FragSpectrumDetails() As udtFragmentationSpectrumDataType       ' 0-based array, ranging from 0 to FragSpectrumIonCount-1

Private FragSpectrumGridLocs() As udtGridLocationType

Private CRCForPredictedSpectrumSaved As Double      ' Used to determine if the values in PredictedSpectrum() have changed
Private CRCForIonListSaved As Double                ' Used to determine if the values in grdIonList() have changed

Private mAutoAligning As Boolean            ' Set to True when auto-aligning (determining the best offset)
Private dblAlignmentScore As Double         ' The score computed for a match
Private lngAlignmentMatchCount As Long      ' # of matches for the current alignment

Private IonMatchList() As Double            ' List of ions to match to those in FragSpectrumDetails()
                                            ' 2D array, 1-based in the first dimension but uses columns 0, 1, and 2
                                            ' Column 0 is the mass, column 1 the raw intensity, column 2 is the normalized intensity (or -1 if not used)
Private IonMatchListCount As Long
Private mIonMatchListDetails As udtIonMatchListDetailsType

Private BinnedData() As Double              ' Binned version of IonMatchList()
                                            ' 2D array, 1-based in the first dimension but uses columns 0, 1, and 2
                                            ' Column 0 is the mass, column 1 the raw intensity, column 2 is the normalized intensity (or -1 if not used)
Private BinnedDataCount As Long             ' When BinnedDataCount = 0, not using BinnedData.  If non-zero, then we are using binned data

Private mDelayUpdate As Boolean
Private mNeedToZoomOutFull As Boolean

Private udtNewFragmentationSpectrumOptions As udtFragmentationSpectrumOptionsType

Private WithEvents objSpectrum As CWSpectrumDLL.Spectrum
Attribute objSpectrum.VB_VarHelpID = -1
'

Private Sub AddIonPairToIonList(strThisLine As String, lngPointsLoaded As Long, blnAllowCommaDelimeter As Boolean, dblMinMass As Double, dblMaxMass As Double)
    Dim strLeftChar As String, strDelimeterList As String
    Dim dblParsedVals(3) As Double      ' 1-based array
    Dim intParseValCount As Integer

    ' Only parse if the first letter (trimmed) is a number (may start with - or +)
    strLeftChar = Left(strThisLine, 1)
    If IsNumeric(strLeftChar) Or strLeftChar = "-" Or strLeftChar = "+" Then
        
        ' Construct Delimeter List: Contains a space, Tab, and possibly comma
        strDelimeterList = " " & vbTab & vbCr & vbLf
        If blnAllowCommaDelimeter Then
            strDelimeterList = strDelimeterList & ","
        End If
 
        intParseValCount = ParseStringValuesDbl(strThisLine, dblParsedVals(), 2, strDelimeterList, , False, False, True)
 
        If intParseValCount >= 2 Then
            
            If dblParsedVals(1) <> 0 And dblParsedVals(2) <> 0 Then
                lngPointsLoaded = lngPointsLoaded + 1
                
                AppendXYPair IonMatchList(), IonMatchListCount, dblParsedVals(1), dblParsedVals(2), True
                
                If dblParsedVals(1) < dblMinMass Then dblMinMass = dblParsedVals(1)
                If dblParsedVals(1) > dblMaxMass Then dblMaxMass = dblParsedVals(1)
            End If
        End If
    End If

End Sub

Private Function AddSpectrumLabels(lngDataCount As Long, dblXVals() As Double, dblYVals() As Double, strDataLabels() As String, intSeriesNumber As Integer) As Long
    ' Returns the number of labels added
    ' Note that strDataLabels() should be 0-based
    
    Dim lngIndex As Long, lngCharIndex As Long, lngLabelsAdded As Long
    Dim strNewCaption As String, strFirstChar As String
    Dim blnShowCaption As Boolean, blnShowAllCaptions As Boolean
    
    If Not (cChkBox(frmIonMatchOptions.chkFragSpecLabelMainIons) Or cChkBox(frmIonMatchOptions.chkFragSpecLabelOtherIons)) Then
        AddSpectrumLabels = 0
        Exit Function
    End If
    lngLabelsAdded = 0
        
    With objSpectrum
        
        .SetAnnotationFontSize intSeriesNumber, ANNOTATION_FONT_SIZE, True
        .SetAnnotationDensityToleranceAutoAdjust True                       ' Could use ToleranceX = 0.9
        
        If cChkBox(frmIonMatchOptions.chkFragSpecLabelMainIons) And cChkBox(frmIonMatchOptions.chkFragSpecLabelOtherIons) Then
            blnShowAllCaptions = True
        Else
            blnShowAllCaptions = False
        End If
        
        For lngIndex = 0 To lngDataCount - 1
            If Len(strDataLabels(lngIndex)) > 0 Then
            
                strNewCaption = strDataLabels(lngIndex)
                
                If blnShowAllCaptions Then
                    blnShowCaption = True
                Else
                    strFirstChar = LCase(Left(strNewCaption, 1))
                    
                    ' See if this is a primary ion label; i.e. a1, b1, y1, a2, b2, y2, etc.
                    blnShowCaption = True
                    If strFirstChar = "a" Or strFirstChar = "b" Or strFirstChar = "y" Then
                        lngCharIndex = 2
                        Do While lngCharIndex <= Len(strNewCaption)
                            If Not IsNumeric(Mid(strNewCaption, lngCharIndex, 1)) Then
                                blnShowCaption = False
                                Exit Do
                            End If
                            lngCharIndex = lngCharIndex + 1
                        Loop
                        
                        If blnShowCaption Then
                            ' This is a primary ion label
                            ' Examine chkFragSpecLabelMainIons
                            blnShowCaption = cChkBox(frmIonMatchOptions.chkFragSpecLabelMainIons)
                        Else
                            blnShowCaption = cChkBox(frmIonMatchOptions.chkFragSpecLabelOtherIons)
                        End If
                    Else
                        blnShowCaption = cChkBox(frmIonMatchOptions.chkFragSpecLabelOtherIons)
                    End If
                End If
            
                ' Do not show captions for shoulder ions
                If InStr(strNewCaption, "Shoulder-") > 0 Then blnShowCaption = False
        
                If blnShowCaption Then
                    
                    .SetAnnotation False, dblXVals(lngIndex) + 1, dblYVals(lngIndex), strNewCaption, 0, intSeriesNumber, asmFixedToSingleDataPoint, lngIndex, False, False, False, True, 0, True
                    
                    lngLabelsAdded = lngLabelsAdded + 1
                End If
            End If
        Next lngIndex
        
    End With
    
    AddSpectrumLabels = lngLabelsAdded
    
End Function

Public Sub AlignmentOffsetValidate()

    If IsNumeric(txtAlignment) Then
        FillIonMatchGridWrapper
        UpdateMassesGridAndSpectrumWrapper
    End If

End Sub

Private Sub AppendXYPair(ByRef ThisSpectrum() As Double, ThisSpectrumCount As Long, dblXData As Double, dblYData As Double, Optional blnIncludeZeroXValue As Boolean = False)
    ' ThisSpectrum() is a 2D array, 1-based in the first column and 0-based in the second column
    ' Note: Will not add data if dblXData = 0 and blnIncludeZeroXValue = False
    
    Dim lngDimCount As Long, lngIndex As Long
    Dim dblSpectrumCopy() As Double         ' 2D array, 1-based in the first dimension, 0 to 2 in the second dimension
    
    If dblXData = 0 And Not blnIncludeZeroXValue Then
        Exit Sub
    End If
    
    lngDimCount = UBound(ThisSpectrum(), 1)
    
    If ThisSpectrumCount >= lngDimCount Then
        ' Need to increase the space in ThisSpectrum()
        Debug.Assert False
        
        ' However, can't use Redim Preserve since ThisSpectrum() is a 2D array
        ReDim dblSpectrumCopy(lngDimCount + 10, 0 To ION_LIST_ARRAY_MAXINDEX_2D)
        For lngIndex = 0 To lngDimCount
            dblSpectrumCopy(lngIndex, ilaMass) = ThisSpectrum(lngIndex, ilaMass)
            dblSpectrumCopy(lngIndex, ilaIntensity) = ThisSpectrum(lngIndex, ilaIntensity)
            dblSpectrumCopy(lngIndex, ilaNormalizedIntensity) = ThisSpectrum(lngIndex, ilaNormalizedIntensity)
        Next lngIndex
        
        ' Now that dblSpectrumCopy() has been populated, we can perform a bulk array copy using the following
        ThisSpectrum() = dblSpectrumCopy()
        
    End If
        
    ThisSpectrumCount = ThisSpectrumCount + 1
    ThisSpectrum(ThisSpectrumCount, ilaMass) = dblXData
    ThisSpectrum(ThisSpectrumCount, ilaIntensity) = dblYData
    
End Sub

Private Sub AutoAlignMatchIonList()
    Dim dblMaxOffset As Double, dblOffsetIncrement As Double
    Dim lngIterationsCompleted As Long, lngPredictedIterationCount As Long
    Dim dblBestScore As Double, dblBestScoreOffset As Double, dblAdditionalBestScoreOffset As Double
    Dim dblCurrentOffset As Double, dblOffsetEnd As Double
    Dim lngSequentialDuplicates As Long
    Dim blnUpdateSpectrumSaved As Boolean
    
    With frmSetValue
        .Caption = LookupMessage(1150)
        .lblStartVal.Caption = LookupMessage(1155)
        .txtStartVal = "10"
        .lblEndVal.Caption = LookupMessage(1160)
        .txtEndVal = 0.1
        
        .SetLimits False
        
        .Show vbModal
    End With
    
    If UCase(frmSetValue.lblHiddenStatus) <> "OK" Then Exit Sub
    
    ' Auto Align the Ion Match List
    With frmSetValue
        If IsNumeric(.txtStartVal) Then
            dblMaxOffset = Abs(CDbl(.txtStartVal))
        Else
            dblMaxOffset = 10
        End If
        
        If dblMaxOffset < 0.5 Or dblMaxOffset > 1000 Then dblMaxOffset = 10
        
        If IsNumeric(.txtEndVal) Then
            dblOffsetIncrement = Abs(CDbl(.txtEndVal))
        Else
            dblOffsetIncrement = 0.1
        End If
        dblOffsetIncrement = Round(dblOffsetIncrement, 3)
        If dblOffsetIncrement < 0.001 Or dblOffsetIncrement > 10 Then dblOffsetIncrement = 0.1
    End With
    
    ' Determine the predicted iteration count
    lngPredictedIterationCount = dblMaxOffset * 2 / dblOffsetIncrement
    
    If lngPredictedIterationCount > 10000 Then
        ' Over 1000 predicted iterations
        ' Raise dblOffsetIncrement so that lngPredictedIterationCount = 10000
        lngPredictedIterationCount = 10000
        dblOffsetIncrement = dblMaxOffset * 2 / lngPredictedIterationCount
    End If
    
    ' Disable automatic plot updating
    blnUpdateSpectrumSaved = mnuUpdateSpectrum.Checked
    mnuUpdateSpectrum.Checked = False

    ' Hide grdIonList and grdFragMasses
    grdIonList.Visible = False
    grdFragMasses.Visible = False
    
    ' Turn on mAutoAligning
    mAutoAligning = True
    
    ' Show the Progress form
    frmProgress.InitializeForm LookupMessage(1165), 0, lngPredictedIterationCount, True
    frmProgress.ToggleAlwaysOnTop True
    
    dblCurrentOffset = -dblMaxOffset
    dblOffsetEnd = dblMaxOffset
    
    lngSequentialDuplicates = 0
    lngIterationsCompleted = 0
    Do While dblCurrentOffset < dblOffsetEnd
        ' Assigning dblCurrentOffset to txtAlignment does not trigger the _Validate function
        txtAlignment = dblCurrentOffset
    
        ' Must set mFragMatchSettingsChanged to True each time before re-matching
        mFragMatchSettingsChanged = True

        ' Update the masses of the match ions
        FillIonMatchGridWrapper

        ' Remove cell coloring from grdFragMasses
        FlexGridRemoveHighlighting grdFragMasses, 1, grdFragMasses.Cols - 2
        
        ' Perform the match
        MatchIons
    
        ' See if new dblMatchScore is better than previous
        ' If so, save new offset as best offset
        If dblAlignmentScore >= dblBestScore Then
            If dblAlignmentScore > dblBestScore Then
                dblBestScore = dblAlignmentScore
                dblBestScoreOffset = dblCurrentOffset
                ' Update Stats since new best
                UpdateMatchCountAndScoreWork lngAlignmentMatchCount, dblAlignmentScore

                ' Set lngSequentialDuplicates
                lngSequentialDuplicates = 1
                ' Reset dblAdditionalBestScoreOffset
                dblAdditionalBestScoreOffset = dblCurrentOffset
            Else
                If dblCurrentOffset - dblOffsetIncrement * 1.51 < dblAdditionalBestScoreOffset Then
                    dblAdditionalBestScoreOffset = dblCurrentOffset
                    lngSequentialDuplicates = lngSequentialDuplicates + 1
                End If
            End If
        End If
        
        ' Update Iterations counter
        lngIterationsCompleted = lngIterationsCompleted + 1
        If lngIterationsCompleted Mod 5 = 0 Then
            frmProgress.UpdateProgressBar lngIterationsCompleted
            If KeyPressAbortProcess > 1 Then
                Exit Do
            End If
        End If
    
        ' Computer new offset
        dblCurrentOffset = Round(dblCurrentOffset + dblOffsetIncrement, 3)
    
    Loop
    KeyPressAbortProcess = 0
    
    If lngSequentialDuplicates > 1 Then
        ' More than one offset in a row gave the same score
        ' dblBestScoreOffset is the first score that matched while dblAdditionalBestScoreOffset is the last
        dblBestScoreOffset = dblBestScoreOffset + Abs((dblAdditionalBestScoreOffset - dblBestScoreOffset)) / 2
    End If
    
'    MsgBox "Best offset is " & dblBestScoreOffset & " (" & lngSequentialDuplicates & " duplicates)", vbInformation + vbOKOnly, "Alignment Results"
    
    ' Hide the progress form
    frmProgress.HideForm
    
    ' Turn off mAutoAligning
    mAutoAligning = False

    ' Store best match in txtAlignment
    ' Does not trigger any match routines, so must manually call
    txtAlignment = dblBestScoreOffset
    
    UpdateIonMatchListWrapper
    
    ' Re-enable mnuUpdateSpectrum
    mnuUpdateSpectrum.Checked = blnUpdateSpectrumSaved

    ' Unhide grdIonList
    grdIonList.Visible = True

End Sub

Private Sub BinIonMatchList()
    
    Const BIN_MEMBER_COUNT_INDEX = 2
    
    Dim lngIndex As Long, lngBinNumber As Long
    Dim lngValuesToPopulate As Long
    Dim lclBinnedData() As Double           ' Binned version of IonMatchList()
                                            ' 2D array, 1-based in the first dimension but uses columns 0, 1, and 2
                                            ' Column 0 is the mass, column 1 the raw intensity, column 2 is the normalized intensity (or -1 if not used)
                                            ' Note, though, that I actually use column 3 to record the number of ions in each bin
                                            
    Dim lclBinnedDataCount As Long          ' When lclBinnedDataCount = 0, not using lclBinnedData.  If non-zero, then we are using binned data
    
    Dim dblMassMinimum As Double, dblMassMaximum As Double
    Dim dblBinWindow As Double
    
    ' Exit sub if user has turned off binning
    If Not cChkBox(frmIonMatchOptions.chkGroupSimilarIons) Or IonMatchListCount = 0 Then
        BinnedDataCount = 0
        ReDim BinnedData(1, 0 To ION_LIST_ARRAY_MAXINDEX_2D)
        Exit Sub
    End If
    
    With frmIonMatchOptions
        dblBinWindow = .txtGroupIonMassWindow
        If dblBinWindow < 0 Or dblBinWindow > 100 Then
            dblBinWindow = 0.5
        End If
    End With

    If dblBinWindow = 0 Then
        ' Window is 0, do not attempt to bin data
        Exit Sub
    End If
    
    ' Find the Mass Limits
    FindMassLimits IonMatchList(), IonMatchListCount, dblMassMinimum, dblMassMaximum

    ' Determine the total number of bins, limiting to MAX_BIN_COUNT if necessary
    lclBinnedDataCount = CLngRoundUp((dblMassMaximum - dblMassMinimum) / dblBinWindow)
    If lclBinnedDataCount > MAX_BIN_COUNT Then
        lclBinnedDataCount = MAX_BIN_COUNT
        dblBinWindow = (dblMassMaximum - dblMassMinimum) / lclBinnedDataCount
    End If
    
    ' Initialize lclBinnedData()
    ReDim lclBinnedData(lclBinnedDataCount + 1, 3)  ' 1-based array in the first dimension
                                                 ' For 2nd dimension, column 0 is mass, column 1 is intensity
                                                 '   and column 2 is # of data points assigned to bin
    ' An explanation on the binning process:
    '  For each mass in IonMatchList() the appropriate bin number is calculated
    ' If the bin is empty, the ion's mass and intensity are assigned to the bin (,0) and (,1)
    '  and a 1 is placed in cell lclBinnedData(lngBinNumber,2)
    ' If the bin already has a value in it, the new mass is added to the current mass
    '  and cell (,2) is incremented by 1
    ' Additionally, the intensities are compared and the higher intensity is assigned to cell (,1)
    ' When all done binning, the values in cell (,0) are divided by the value in (,2) to get an
    '  averag mass value for all of the masses assigned to the bin
    
    ' Step through IonMatchList() and assign data to appropriate bin
    For lngIndex = 1 To IonMatchListCount
        lngBinNumber = MassToBinNumber(IonMatchList(lngIndex, ilaMass), dblMassMinimum, dblBinWindow)
        If lclBinnedData(lngBinNumber, ilaMass) = 0 Then
            ' No ions assigned to this bin yet; assign mass and intensity
            lclBinnedData(lngBinNumber, ilaMass) = IonMatchList(lngIndex, ilaMass)
            lclBinnedData(lngBinNumber, ilaIntensity) = IonMatchList(lngIndex, ilaIntensity)
            lclBinnedData(lngBinNumber, BIN_MEMBER_COUNT_INDEX) = 1
        Else
            lclBinnedData(lngBinNumber, ilaMass) = lclBinnedData(lngBinNumber, ilaMass) + IonMatchList(lngIndex, ilaMass)
            If IonMatchList(lngIndex, ilaIntensity) > lclBinnedData(lngBinNumber, ilaIntensity) Then
                lclBinnedData(lngBinNumber, ilaIntensity) = IonMatchList(lngIndex, ilaIntensity)
            End If
            lclBinnedData(lngBinNumber, BIN_MEMBER_COUNT_INDEX) = lclBinnedData(lngBinNumber, BIN_MEMBER_COUNT_INDEX) + 1
        End If
    Next lngIndex
    
    ' Step through lclBinnedData() and compute the correct mass
    ' Additionally, determine the number of bins with actual data in them
    lngValuesToPopulate = 0
    For lngBinNumber = 1 To lclBinnedDataCount
        If lclBinnedData(lngBinNumber, BIN_MEMBER_COUNT_INDEX) > 0 Then
            lngValuesToPopulate = lngValuesToPopulate + 1
            lclBinnedData(lngBinNumber, ilaMass) = lclBinnedData(lngBinNumber, ilaMass) / lclBinnedData(lngBinNumber, BIN_MEMBER_COUNT_INDEX)
        End If
    Next lngBinNumber
    
    ' Copy data from lclBinnedData() to BinnedData()
    ' Only copy non-zero values
    BinnedDataCount = lngValuesToPopulate
    
    
    ' Redimension BinnedData() based on lngValuesToPopulate
    ReDim BinnedData(BinnedDataCount + 1, 0 To ION_LIST_ARRAY_MAXINDEX_2D)      ' 1-based array in the first dimension
                                                                                ' For 2nd dimension, column 0 is mass, column 1 is intensity
                                                                                '   and column 2 is the normalized intensity
    
    ' Load data into BinnedData()
    lngIndex = 0
    For lngBinNumber = 1 To lclBinnedDataCount
        If lclBinnedData(lngBinNumber, BIN_MEMBER_COUNT_INDEX) > 0 Then
            lngIndex = lngIndex + 1
            BinnedData(lngIndex, ilaMass) = lclBinnedData(lngBinNumber, ilaMass)
            BinnedData(lngIndex, ilaIntensity) = lclBinnedData(lngBinNumber, ilaIntensity)
            BinnedData(lngIndex, ilaNormalizedIntensity) = BinnedData(lngIndex, ilaIntensity)
        End If
    Next lngBinNumber
    
End Sub

Private Sub CheckSequenceTerminii()
    ' If 3 letter codes are enabled, then checks to see if the sequence begins with H and ends with OH
    ' If so, makes sure the first three letters are not an amino acid
    ' If they're not, removes the H and OH and sets cboNTerminus and cboCTerminus accordingly
    
    Dim strCleanSequence As String
    Dim lngAbbrevID  As Long, strSymbol As String
    Dim blnIsAminoAcid As Boolean
    
    mDelayUpdate = True
    
    If Len(txtSequence) > 0 And Get3LetterCodeState = True Then
        If UCase(Left(txtSequence, 1)) = "H" Then
            If UCase(Right(txtSequence, 2)) = "OH" Then
                If IsCharacter(Mid(txtSequence, 2, 1)) And IsCharacter(Mid(txtSequence, 3, 1)) Then
                    lngAbbrevID = objMwtWin.GetAbbreviationID(Left(txtSequence, 3))
                    If lngAbbrevID > 0 Then
                        ' Matched an abbreviation; is it an amino acid?
                        objMwtWin.GetAbbreviation lngAbbrevID, strSymbol, "", 0, blnIsAminoAcid
                        If Not blnIsAminoAcid Then
                            ' Matched an abbreviation, but it's not an amino acid
                            lngAbbrevID = 0
                        End If
                    End If
                Else
                    lngAbbrevID = 0
                End If
                
                If lngAbbrevID = 0 Then
                    ' The first three characters do not represent a 3 letter amino acid code
                    ' Remove the H and OH
                    strCleanSequence = Mid(txtSequence, 2, Len(txtSequence) - 3)
                    If Left(strCleanSequence, 1) = "-" Then strCleanSequence = Mid(strCleanSequence, 2)
                    If Right(strCleanSequence, 1) = "-" Then strCleanSequence = Left(strCleanSequence, Len(strCleanSequence) - 1)
                    txtSequence = strCleanSequence
                    
                    cboNTerminus.ListIndex = ntgHydrogen
                    cboCTerminus.ListIndex = ctgHydroxyl
                End If
            End If
        End If
    End If
    
    mDelayUpdate = False

End Sub

Private Sub ConvertSequenceMH(blnFavorMH As Boolean)
    ' When blnFavorMH = True then computes MHAlt using MH
    ' Otherwise, computes MH using MHAlt
    
    Static blnWorking As Boolean
    Dim dblMH As Double, dblMHAlt As Double, intCharge As Integer
    
    If blnWorking Then Exit Sub
    blnWorking = True
    
    intCharge = cboMHAlternate.ListIndex + 1
    If intCharge < 1 Then intCharge = 1
    
    If blnFavorMH Then
        dblMH = CDblSafe(txtMH)
        If dblMH > 0 Then
            dblMHAlt = objMwtWin.ConvoluteMass(dblMH, 1, intCharge)
        Else
            dblMHAlt = 0
        End If
        txtMHAlt = Round(dblMHAlt, 6)
    Else
        dblMHAlt = CDblSafe(txtMHAlt)
        If dblMHAlt > 0 Then
            dblMH = objMwtWin.ConvoluteMass(dblMHAlt, intCharge, 1)
        Else
            dblMH = 0
        End If
        txtMH = Round(dblMH, 6)
    End If
    
    blnWorking = False
    
End Sub

Private Sub CopyFragGridInfo(grdThisGrid As MSFlexGrid, Optional eCopyMode As gcmGridCopyModeConstants = gcmText, Optional intRowStart As Integer = -1, Optional intRowEnd As Integer = -1, Optional intColStart As Integer = -1, Optional intColEnd As Integer = -1)
    ' intCopyMode can be 0 = text, 1 = RTF, or 2 = Html
    
    If grdThisGrid.Rows > 1 Then
        If intRowStart < 0 Then intRowStart = 0
        If intRowEnd < 0 Then intRowEnd = grdThisGrid.Rows - 1
        If intColStart < 0 Then intColStart = 0
        If intColEnd < 0 Then intColEnd = grdThisGrid.Cols - 1
        
        FlexGridCopyInfo Me, grdThisGrid, eCopyMode, intRowStart, intRowEnd, intColStart, intColEnd
    End If
    
End Sub
    
Private Sub CopySequenceMW()
    Dim strWork As String
    
    strWork = Trim(objMwtWin.Peptide.GetPeptideMass())
    
    Clipboard.SetText strWork, vbCFText

End Sub

Public Sub DisplaySpectra()
    Dim lngIndex As Long
    
    Static blnSeriesPlotModeInitialized(MAX_SERIES_COUNT) As Boolean
    
    Dim PredictedSpectrumCRC As Double
    Dim dblMassTimesIntensity As Double
    Dim blnUpdateFragSpectrum As Boolean, blnUpdateIonListSpectrum As Boolean
    Dim blnAllowResumeNextErrorHandling As Boolean
    Dim blnEmphasizeProlineIons As Boolean
    
    Dim dblXVals() As Double            ' 0-based array
    Dim dblYVals() As Double            ' 0-based array
    Dim strDataLabels() As String       ' 0-based array
    Dim lngDataCount As Long
    Dim intSeriesNumber As Integer
    
    Dim blnShowAnnotations As Boolean, blnRemoveExistingAnnotations As Boolean
    Dim blnCursorVisibilitySaved As Boolean
    Dim blnAutoHideCaptionsSaved As Boolean
    
    If Not mFormActivatedByUser Then Exit Sub
    
On Error GoTo DisplaySpectraErrorHandler
    blnAllowResumeNextErrorHandling = True
    
    ' Compute a CRC value for the current PredictedSpectrum() array and compare to the previously computed value
    PredictedSpectrumCRC = 0
    For lngIndex = 0 To FragSpectrumIonCount - 1
        dblMassTimesIntensity = Abs((lngIndex + 1) * FragSpectrumDetails(lngIndex).Mass * FragSpectrumDetails(lngIndex).Intensity)
        If dblMassTimesIntensity > 0 Then
            PredictedSpectrumCRC = PredictedSpectrumCRC + Log(dblMassTimesIntensity)
        End If
    Next lngIndex
    
    ' If the new CRC is different than the old one then re-plot the spectrun
    If PredictedSpectrumCRC <> CRCForPredictedSpectrumSaved Then
        CRCForPredictedSpectrumSaved = PredictedSpectrumCRC
        blnUpdateFragSpectrum = True
    End If
    
    ' Also compute a CRC value for the ions in grdIonList and compare to the previously computed value
    PredictedSpectrumCRC = 0
    With grdIonList
        lngDataCount = .Rows - 1
        For lngIndex = 1 To lngDataCount
            dblMassTimesIntensity = Abs(lngIndex * .TextMatrix(lngIndex, ilgMass) * .TextMatrix(lngIndex, ilgIntensity))
            If dblMassTimesIntensity > 0 Then
                PredictedSpectrumCRC = PredictedSpectrumCRC + Log(dblMassTimesIntensity)
            End If
        Next lngIndex
    End With
    
    ' If the new CRC is different than the old one then re-plot the spectrun
    If PredictedSpectrumCRC <> CRCForIonListSaved Then
        CRCForIonListSaved = PredictedSpectrumCRC
        blnUpdateIonListSpectrum = True
    End If
    
    If Not (blnUpdateFragSpectrum Or blnUpdateIonListSpectrum) Then
        Exit Sub
    End If
    
    blnAllowResumeNextErrorHandling = False
    
    objSpectrum.ShowSpectrum
    
    ' Hide the Cursor and disable auto-hiding of annotations
    blnCursorVisibilitySaved = objSpectrum.GetCursorVisibility()
    If blnCursorVisibilitySaved Then objSpectrum.SetCursorVisible False

    blnAutoHideCaptionsSaved = objSpectrum.GetAnnotationDensityAutoHideCaptions()
    If blnAutoHideCaptionsSaved Then objSpectrum.SetAnnotationDensityAutoHideCaptions False, False
    
    blnEmphasizeProlineIons = cChkBox(frmIonMatchOptions.chkFragSpecEmphasizeProlineYIons)
    
    lngDataCount = FragSpectrumIonCount
    If lngDataCount > 0 Then
        ReDim dblXVals(0 To lngDataCount - 1)
        ReDim dblYVals(0 To lngDataCount - 1)
        ReDim strDataLabels(0 To lngDataCount - 1)
    Else
        ReDim dblXVals(0)
        ReDim dblYVals(0)
        ReDim strDataLabels(0)
    End If
    
    ' Now fill dblXVals(), dblYVals(), and strDataLabels()
    For lngIndex = 0 To lngDataCount - 1
        With FragSpectrumDetails(lngIndex)
            dblXVals(lngIndex) = .Mass
            dblYVals(lngIndex) = .Intensity
            strDataLabels(lngIndex) = .Symbol
            
            If blnEmphasizeProlineIons Then
                ' Adjust the intensity of all of the peaks to 90% of their original value
                ' Do not adjust the intensity for any proline y ions, thus leaving them at the original intensity
                If LCase(.SymbolGeneric) = "y" Then
                    If .SourceResidueSymbol3Letter = "Pro" Then
                        ' Leave the intensity untouched
                    Else
                        dblYVals(lngIndex) = dblYVals(lngIndex) * 0.9
                    End If
                Else
                    dblYVals(lngIndex) = dblYVals(lngIndex) * 0.9
                End If
            End If
        End With
    Next lngIndex
    
    ' Display the values in the mass spectrum
    ' Only label the predicted spectrum if it is plotted inverted or if grdIonList.Rows() <=1
    intSeriesNumber = 1
    blnShowAnnotations = (cChkBox(frmIonMatchOptions.chkPlotSpectrumInverted) Or grdIonList.Rows <= 1)
    blnRemoveExistingAnnotations = True
    DisplaySpectraWork objSpectrum, intSeriesNumber, lngDataCount, dblXVals(), dblYVals(), strDataLabels(), True, "Theoretical", blnSeriesPlotModeInitialized(), 0, blnShowAnnotations, blnRemoveExistingAnnotations
    
    ' Also plot ions in grdIonList() if present
    If grdIonList.Rows > 1 Then
        With grdIonList
            lngDataCount = .Rows - 1
            ReDim dblXVals(0 To lngDataCount - 1)
            ReDim dblYVals(0 To lngDataCount - 1)
            ReDim strDataLabels(0 To lngDataCount - 1)
            
            For lngIndex = 1 To lngDataCount
                dblXVals(lngIndex - 1) = .TextMatrix(lngIndex, ilgMass)
                dblYVals(lngIndex - 1) = .TextMatrix(lngIndex, ilgIntensity)
                strDataLabels(lngIndex - 1) = .TextMatrix(lngIndex, ilgSymbolMatch)
            Next lngIndex
        End With
        
        ' Display the values in the mass spectrum
        intSeriesNumber = 2
        blnShowAnnotations = True
        blnRemoveExistingAnnotations = Not cChkBox(frmIonMatchOptions.chkAutoLabelMass)     ' Do not remove existing annotations if Auto-Label Mass is True
        DisplaySpectraWork objSpectrum, intSeriesNumber, lngDataCount, dblXVals(), dblYVals(), strDataLabels(), False, mIonMatchListDetails.Caption, blnSeriesPlotModeInitialized(), mIonMatchListDetails.OriginalMaximumIntensity, blnShowAnnotations, blnRemoveExistingAnnotations
    End If
    
    ' The following will change the active series number to 2 if data was present in grdIonList
    ' Otherwise, it will be set to 1
    objSpectrum.SetSpectrumFormCurrentSeriesNumber intSeriesNumber
    
    With objSpectrum
        If mNeedToZoomOutFull Then
            .ZoomOutFull
            mNeedToZoomOutFull = False
        End If
    
        .SetCursorPosition 100, 0
    
        ' Auto-hiding of captions was disabled above; re-enable if necessary
        .SetAnnotationDensityAutoHideCaptions blnAutoHideCaptionsSaved

        ' The cursor was hidden above; re-show it if necessary
        If blnCursorVisibilitySaved Then
            .SetCursorVisible True
        End If
        
        ' Make sure the tick spacing is set to the default
        .SetCustomTicksXAxis 0, 0, True
        .SetCustomTicksYAxis 0, 0, True
        
        .ShowSpectrum
    End With
    
    ' Return the focus back to this form (if possible)
    On Error Resume Next
    Me.SetFocus
    Exit Sub

DisplaySpectraErrorHandler:
    If blnAllowResumeNextErrorHandling Then
        ' Something is probably wrong with the CRC computation routines
        Debug.Assert False
        Resume Next
    Else
        GeneralErrorHandler "frmFragmentationModelling|DisplaySpectra", Err.Number, Err.Description
    End If

End Sub

Private Sub DisplaySpectraWork(objThisCWSpectrum As CWSpectrumDLL.Spectrum, intSeriesNumber As Integer, lngDataCount As Long, dblXVals() As Double, dblYVals() As Double, strDataLabels() As String, blnFragmentationData As Boolean, strLegendCaption As String, ByRef blnSeriesPlotModeInitialized() As Boolean, Optional dblOriginalMaximumIntensity As Double = 0, Optional blnDisplayAnnotations As Boolean = True, Optional blnRemoveExistingAnnotations As Boolean = True)
    
    Dim udtAutoLabelPeaksSettings As udtAutoLabelPeaksOptionsType
    
    With objThisCWSpectrum
        If .GetSeriesCount() < intSeriesNumber Then
            .SetSeriesCount intSeriesNumber
        End If
        
        If Not blnSeriesPlotModeInitialized(intSeriesNumber) Then
            .SetSeriesPlotMode intSeriesNumber, pmStickToZero, False
            blnSeriesPlotModeInitialized(intSeriesNumber) = True
        End If
        
        If lngDataCount <= 0 Then
            .ClearData intSeriesNumber
        Else
            .SetDataXvsY intSeriesNumber, dblXVals(), dblYVals(), lngDataCount, strLegendCaption, dblOriginalMaximumIntensity
        
            If blnFragmentationData Then
                ' Fragmentation MS data
                .SetSeriesColor intSeriesNumber, frmIonMatchOptions.lblFragSpectrumColor.BackColor
            Else
                ' Normal data
                .SetSeriesColor intSeriesNumber, frmIonMatchOptions.lblMatchingIonDataColor.BackColor
                If cChkBox(frmIonMatchOptions.chkAutoLabelMass) Then
                    udtAutoLabelPeaksSettings = .GetAutoLabelPeaksOptions()
                    
                    .AutoLabelPeaks intSeriesNumber, udtAutoLabelPeaksSettings.DisplayXPos, udtAutoLabelPeaksSettings.DisplayYPos, udtAutoLabelPeaksSettings.CaptionAngle, udtAutoLabelPeaksSettings.IncludeArrow, udtAutoLabelPeaksSettings.HideInDenseRegions, udtAutoLabelPeaksSettings.PeakLabelCountMaximum, False, True
                End If
            End If
            
            If blnDisplayAnnotations Then
                If blnRemoveExistingAnnotations Then
                    objSpectrum.RemoveAnnotationsForSeries intSeriesNumber
                End If
                
                AddSpectrumLabels lngDataCount, dblXVals(), dblYVals(), strDataLabels(), intSeriesNumber
            End If
        End If
    End With

End Sub

Private Sub DisplayPredictedIonMasses()
    ' Call objMwtWin to get the predicted fragmentation spectrum masses and intensities
    ' Use this data to populate grdFragMasses
    
    Dim strColumnHeaders() As String
    Dim lngColCount As Long
    
    Dim lngIonIndex As Long
    Dim lngResidueIndex As Long, lngResidueCount As Long
    Dim strSymbol As String, strSymbol1Letter As String
    Dim dblResidueMass As Double
    
    Dim blnIsModified As Boolean
    Dim lngModIDs() As Long
    Dim intModCount As Integer, intModIndex As Integer
    Dim strModSymbol As String, strModSymbolForThisResidue As String
    
    Dim blnUse3LetterSymbol As Boolean
    Dim lngErrorID As Long
    Dim lngColIndex As Long, lngSeqColIndex As Long
    Dim lngThisRow As Long
    Dim strSymbolGeneric As String
    
On Error GoTo DisplayPredictedIonMassesErrorHandler

    UpdateFragmentationSpectrumOptions
    
    ' The GetFragmentationMasses() function computes the masses, intensities, and symbols for the given sequence
    FragSpectrumIonCount = objMwtWin.Peptide.GetFragmentationMasses(FragSpectrumDetails())
    
    ' Initialize (and clear) grdFragMasses
    grdFragMasses.Visible = False
    lngSeqColIndex = InitializeFragMassGrid(strColumnHeaders(), lngColCount)
    
    ' Now populate grdFragMasses with the data
    lngResidueCount = objMwtWin.Peptide.GetResidueCount
    blnUse3LetterSymbol = Get3LetterCodeState()
    For lngResidueIndex = 1 To lngResidueCount

        lngErrorID = objMwtWin.Peptide.GetResidue(lngResidueIndex, strSymbol, dblResidueMass, blnIsModified, intModCount)
        
        strModSymbolForThisResidue = ""
        If blnIsModified Then
            ' I really shouldn't have to ReDim lngModIDs here, but the Dll is refusing to do it, thus generating an error
            ReDim lngModIDs(MAX_MODIFICATIONS)
            
            intModCount = objMwtWin.Peptide.GetResidueModificationIDs(lngResidueIndex, lngModIDs())
            
            For intModIndex = 1 To intModCount
                objMwtWin.Peptide.GetModificationSymbol lngModIDs(intModIndex), strModSymbol, 0, 0, ""
                strModSymbolForThisResidue = strModSymbolForThisResidue & strModSymbol
            Next intModIndex
        End If
        
        Debug.Assert lngErrorID = 0

        ' Add a new row to grdFragMasses
        With grdFragMasses
            .Rows = .Rows + 1
            lngThisRow = .Rows - 1

            ' Add residue number to first column
            .TextMatrix(lngThisRow, 0) = lngResidueIndex

            ' Add reverse residue number to last column
            .TextMatrix(lngThisRow, .Cols - 1) = lngResidueCount - lngResidueIndex + 1

            ' Add the immonium ion to the second column
            .TextMatrix(lngThisRow, 1) = CStrIfNonZero(objMwtWin.Peptide.ComputeImmoniumMass(dblResidueMass), "", 2)

            ' Add the residue symbol
            If blnUse3LetterSymbol Then
                .TextMatrix(lngThisRow, lngSeqColIndex) = strSymbol & strModSymbolForThisResidue
            Else
                strSymbol1Letter = objMwtWin.GetAminoAcidSymbolConversion(strSymbol, False)
                If Len(strSymbol1Letter) = 0 Then strSymbol1Letter = strSymbol
                If strSymbol1Letter = "Xxx" Then strSymbol1Letter = "X"
                .TextMatrix(lngThisRow, lngSeqColIndex) = strSymbol1Letter & strModSymbolForThisResidue
            End If
        End With

    Next lngResidueIndex
    grdFragMasses.Visible = True
    
    ' Initialize FragSpectrumGridLocs()
    ReDim FragSpectrumGridLocs(FragSpectrumIonCount)
    
    ' Finally, step through FragSpectrumDetails and populate grdFragMasses with the ion masses
    '  Shoulder ion masses are not displayed, but FragSpectrumGridLocs is updated with the associated primary ion
    For lngIonIndex = 0 To FragSpectrumIonCount - 1
        With FragSpectrumDetails(lngIonIndex)
            ' Can start at column 2 since column 0 is # and column 1 is Immon.
            ' Can stop at grdFragMasses.Cols - 2 since the final column is #
            
            strSymbolGeneric = .SymbolGeneric
            If .IsShoulderIon Then
                strSymbolGeneric = Replace(strSymbolGeneric, SHOULDER_ION_PREFIX, "")
            End If
            
            For lngColIndex = 2 To grdFragMasses.Cols - 2
                If grdFragMasses.TextMatrix(0, lngColIndex) = strSymbolGeneric Then
                    ' Add the mass of this ion to the appropriate row in this column
                    If .SourceResidueNumber < grdFragMasses.Rows Then
                        If Not .IsShoulderIon Then
                            ' Only display if not a Shoulder Ion
                            grdFragMasses.TextMatrix(.SourceResidueNumber, lngColIndex) = CStrIfNonZero(.Mass, "", 2)
                        End If
                        
                        ' Update FragSpectrumGridLocs() with the Row Index and Column Index
                        FragSpectrumGridLocs(lngIonIndex).Row = .SourceResidueNumber
                        FragSpectrumGridLocs(lngIonIndex).Col = lngColIndex
                    Else
                        ' Invalid residue number; this is unexpected
                        Debug.Assert False
                    End If
                End If
            Next lngColIndex
        End With
    Next lngIonIndex
    
    Exit Sub
    
DisplayPredictedIonMassesErrorHandler:
    GeneralErrorHandler "frmFragmentationModelling|DisplayPredictedIonMasses", Err.Number, Err.Description
    
End Sub


Private Sub EnableDisableControls()
    Dim boolRemovePrecursorIon As Boolean
    
    cboDoubleCharge.Enabled = cChkBox(chkDoubleCharge)
    cboTripleCharge.Enabled = cChkBox(chkTripleCharge)
    lblChargeThreshold.Enabled = cboDoubleCharge.Enabled Or cboTripleCharge.Enabled
    
    boolRemovePrecursorIon = cChkBox(chkRemovePrecursorIon)
    lblPrecursorIonMass.Enabled = boolRemovePrecursorIon
    txtPrecursorIonMass.Enabled = boolRemovePrecursorIon
    lblPrecursorMassWindow.Enabled = boolRemovePrecursorIon
    txtPrecursorMassWindow.Enabled = boolRemovePrecursorIon
    lblDaltons(0).Enabled = boolRemovePrecursorIon
    lblDaltons(1).Enabled = boolRemovePrecursorIon
        
End Sub

Private Sub InitializeDummyData(intDataType As Integer)
    ' intDataType can be 0: continuous sine wave
    '                    1: stick data (only 20 points)
    '                    2: stick data (1000's of points, mostly zero, with a few spikes)

    Dim ThisXYDataSet As usrXYDataSet
    Dim lngIndex As Long, sngOffset As Single
    
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
    Case 2
        With ThisXYDataSet
            .XYDataListCount = 50000
            ReDim .XYDataList(.XYDataListCount)
            For lngIndex = 1 To .XYDataListCount
                .XYDataList(lngIndex).XVal = 100 + lngIndex / 500
                If lngIndex Mod 5000 = 0 Then
                    .XYDataList(lngIndex).YVal = Rnd(1) * .XYDataListCount / 200 * Rnd(1)
                ElseIf lngIndex Mod 3000 = 0 Then
                    .XYDataList(lngIndex).YVal = Rnd(1) * .XYDataListCount / 650 * Rnd(1)
                Else
                    .XYDataList(lngIndex).YVal = Rnd(1) * 3
                End If
            Next lngIndex
        End With
    Case Else
        With ThisXYDataSet
            .XYDataListCount = 360! * 100!
            
            ReDim .XYDataList(.XYDataListCount)
            sngOffset = 10
            For lngIndex = 1 To .XYDataListCount
                If lngIndex Mod 5050 = 0 Then
                    sngOffset = Rnd(1) + 10
                End If
                .XYDataList(lngIndex).XVal = CDbl(lngIndex) / 1000 - 5
                .XYDataList(lngIndex).YVal = sngOffset - Abs((lngIndex - .XYDataListCount / 2)) / 10000 + Sin(DegToRadiansMultiplier * lngIndex) * Cos(DegToRadiansMultiplier * lngIndex / 2) * 1.29967878493163
            Next lngIndex
        End With
    End Select
    
    ' Fill Ion IonMatchList
    InitializeIonMatchList ThisXYDataSet.XYDataListCount
    
    With ThisXYDataSet
        For lngIndex = 1 To .XYDataListCount
            IonMatchList(lngIndex, ilaMass) = .XYDataList(lngIndex).XVal
            IonMatchList(lngIndex, ilaIntensity) = .XYDataList(lngIndex).YVal
        Next lngIndex
        IonMatchListCount = .XYDataListCount
    End With
    
    UpdateIonMatchListWrapper
        
End Sub

Private Sub InitializeIonListGrid()
    
    With grdIonList
        .Clear
        .Rows = 1
        .Cols = ION_LIST_COL_COUNT
        .TextMatrix(0, ilgMass) = LookupLanguageCaption(12550, "Mass")
        .TextMatrix(0, ilgIntensity) = LookupLanguageCaption(12560, "Intensity")
        .TextMatrix(0, ilgSymbolMatch) = LookupLanguageCaption(12570, "Symbol")
        
        .ColWidth(ilgMass) = 900
        .ColWidth(ilgIntensity) = 1000
        .ColWidth(ilgSymbolMatch) = 1000
        .ColAlignment(ilgMass) = flexAlignRightCenter
        .ColAlignment(ilgIntensity) = flexAlignRightCenter
        .ColAlignment(ilgSymbolMatch) = flexAlignCenterCenter
    End With
End Sub
    
Private Function InitializeFragMassGrid(strColumnHeaders() As String, lngColCount As Long) As Long
    ' Initializes grdFragMasses using the headers in strColumnHeaders() and the widths in lngColumnWidths()
    ' Returns the column number that will hold the residue symbol (lngSeqColIndex)
    '  this column is located just before the first y ion column
    
    Dim lngColumnWidths() As Long
    Dim lngColIndex As Long
    Dim lngSeqColIndex As Long
    Dim lngIonIndex As Long, lngIndexCompare As Long
    Dim blnMatched As Boolean, blnSeqColumnAdded As Boolean
        
    Dim strYIonSymbol As String
    Dim strZIonSymbol As String
    Dim strFirstChar As String
    
    Dim strColumnHeadersToAdd() As String
    Dim lngColHeadersToAddCount As Long
    Dim lngColHeadersToAddDimCount As Long
    
On Error GoTo InitializeFragMassGridErrorHandler

    lngColHeadersToAddCount = 0
    lngColHeadersToAddDimCount = 10
    ReDim strColumnHeadersToAdd(lngColHeadersToAddDimCount)
    
    ' Examine FragSpectrumDetails().Symbol to make a list of the possible ion types present
    For lngIonIndex = 0 To FragSpectrumIonCount - 1
        With FragSpectrumDetails(lngIonIndex)
            If Not .IsShoulderIon Then
                blnMatched = False
                For lngIndexCompare = 0 To lngColHeadersToAddCount - 1
                    If strColumnHeadersToAdd(lngIndexCompare) = .SymbolGeneric Then
                        blnMatched = True
                        Exit For
                    End If
                Next lngIndexCompare
                
                If Not blnMatched Then
                    strColumnHeadersToAdd(lngColHeadersToAddCount) = .SymbolGeneric
                    
                    ' Increment lngColHeadersToAddCount and reserve more memory if needed
                    lngColHeadersToAddCount = lngColHeadersToAddCount + 1
                    If lngColHeadersToAddCount >= lngColHeadersToAddDimCount Then
                        lngColHeadersToAddDimCount = lngColHeadersToAddDimCount + 10
                        ReDim Preserve strColumnHeadersToAdd(lngColHeadersToAddDimCount)
                    End If
                End If
            End If
        End With
    Next lngIonIndex
    
    ' There are at a minimum 4 columns (#, Immon., Seq., and #)
    ' We'll start with # and Immon.
    ReDim strColumnHeaders(4 + lngColHeadersToAddCount)
    ReDim lngColumnWidths(4 + lngColHeadersToAddCount)
    
    lngColCount = 2
    strColumnHeaders(0) = LookupLanguageCaption(12500, "#")
    strColumnHeaders(1) = LookupLanguageCaption(12510, "Immon.")
    
    lngColumnWidths(0) = 400
    lngColumnWidths(1) = 700
    
    strYIonSymbol = objMwtWin.Peptide.LookupIonTypeString(itYIon)
    strZIonSymbol = objMwtWin.Peptide.LookupIonTypeString(itZIon)
    
    If lngColHeadersToAddCount > 0 Then
        ' Sort strColumnHeadersToAdd() alphabetically (a, b, y, c, or z)
        ShellSortString strColumnHeadersToAdd(), 0, lngColHeadersToAddCount - 1
    
        ' Append the items in strColumnHeadersToAdd() to strColumnHeaders
        For lngColIndex = 0 To lngColHeadersToAddCount - 1
            If lngColIndex < lngColHeadersToAddCount And Not blnSeqColumnAdded Then
                ' Check if this column is the first y-ion or z-ion column
                strFirstChar = Left(strColumnHeadersToAdd(lngColIndex), 1)
                If strFirstChar = strYIonSymbol Or strFirstChar = strZIonSymbol Then
                    strColumnHeaders(lngColCount) = LookupLanguageCaption(12520, "Seq.")
                    lngSeqColIndex = lngColCount
                    lngColCount = lngColCount + 1
                    blnSeqColumnAdded = True
                End If
            End If
            
            strColumnHeaders(lngColCount) = strColumnHeadersToAdd(lngColIndex)
            lngColumnWidths(lngColCount) = 800
            lngColCount = lngColCount + 1
        
        Next lngColIndex
    
    End If
    
    ' If the sequence column still wasn't added, then add it now
    If Not blnSeqColumnAdded Then
        strColumnHeaders(lngColCount) = LookupLanguageCaption(12520, "Seq.")
        lngSeqColIndex = lngColCount
        lngColCount = lngColCount + 1
    End If
    lngColumnWidths(lngSeqColIndex) = 600

    ' The final column is a duplicate of the zeroth column
    strColumnHeaders(lngColCount) = strColumnHeaders(0)
    lngColumnWidths(lngColCount) = lngColumnWidths(0)
    lngColCount = lngColCount + 1
    
    ' Now populate the grid
    With grdFragMasses
        .Clear
        .Rows = 1
        .Cols = lngColCount
        
        For lngColIndex = 0 To lngColCount - 1
            .TextMatrix(0, lngColIndex) = strColumnHeaders(lngColIndex)
            .ColWidth(lngColIndex) = lngColumnWidths(lngColIndex)
            .ColAlignment(lngColIndex) = flexAlignCenterCenter
        Next lngColIndex
    
    End With

    InitializeFragMassGrid = lngSeqColIndex
    Exit Function
    
InitializeFragMassGridErrorHandler:
    GeneralErrorHandler "frmFragmentationModelling|InitializeFragMassGrid", Err.Number, Err.Description
    
End Function

Private Sub InitializeIonMatchList(ByVal lngNumberPoints As Long)
    ' Redimension array (erasing old values) using size lngNumberPoints
    
    IonMatchListCount = 0
    ReDim IonMatchList(lngNumberPoints, 0 To ION_LIST_ARRAY_MAXINDEX_2D)  ' 1-based array in the first dimension, uses columns 0, 1, and 2 in the 2nd dimension
        
    UpdateIonLoadStats True
    
End Sub

Private Sub FillIonMatchGridWrapper()
    If BinnedDataCount > 0 Then
        FillIonMatchGrid BinnedData(), BinnedDataCount
    Else
        FillIonMatchGrid IonMatchList(), IonMatchListCount
    End If
End Sub

Private Sub FillIonMatchGrid(ThisIonList() As Double, ThisIonListCount As Long)
    ' ThisIonList is a 2D array, 1-based in the first dimension but uses columns 0, 1, and 2
    ' Column 0 is the mass, column 1 the raw intensity, column 2 is the normalized intensity (or -1 if not used)
    
    Dim lngIndex As Long, lngValuesToPopulate As Long, dblAlignmentValue As Double
    
    InitializeIonListGrid
    
    lngValuesToPopulate = 0
    For lngIndex = 1 To ThisIonListCount
        If ThisIonList(lngIndex, ilaNormalizedIntensity) >= 0 Then
            lngValuesToPopulate = lngValuesToPopulate + 1
        End If
    Next lngIndex
    
    ' The Alignment Value is added to the masses of all ions
    dblAlignmentValue = CDblSafe(txtAlignment)
    
    With grdIonList
        .Rows = lngValuesToPopulate + 1
        lngValuesToPopulate = 0
        For lngIndex = 1 To ThisIonListCount
            If ThisIonList(lngIndex, ilaNormalizedIntensity) >= 0 Then
                lngValuesToPopulate = lngValuesToPopulate + 1
                .TextMatrix(lngValuesToPopulate, ilgMass) = Format(ThisIonList(lngIndex, ilaMass) + dblAlignmentValue, "#0.00")
                .TextMatrix(lngValuesToPopulate, ilgIntensity) = Format(ThisIonList(lngIndex, ilaNormalizedIntensity), "#0.00")
                .TextMatrix(lngValuesToPopulate, ilgSymbolMatch) = ""
            End If
        Next lngIndex
    End With
End Sub

Private Sub FindMassLimits(ThisIonMatchList() As Double, ThisIonMatchListCount As Long, dblMassMinimum As Double, dblMassMaximum As Double)
    Dim lngIndex As Long, dblMassValue As Double
    
    ' Determine the smallest and largest ion mass (m/z)
    dblMassMinimum = HighestValueForDoubleDataType
    dblMassMaximum = LowestValueForDoubleDataType
    For lngIndex = 1 To ThisIonMatchListCount
        dblMassValue = ThisIonMatchList(lngIndex, ilaMass)
        If dblMassValue > dblMassMaximum Then dblMassMaximum = dblMassValue
        If dblMassValue < dblMassMinimum Then dblMassMinimum = dblMassValue
    Next lngIndex

End Sub

Private Sub FlexGridKeyPressHandler(frmThisForm As VB.Form, grdThisGrid As MSFlexGrid, KeyCode As Integer, Shift As Integer, Optional lngMaxColumnToSelect As Long = -1)
    If KeyCode = vbKeyC And (Shift And vbCtrlMask) Then
        ' Ctrl+C
        FlexGridCopyInfo frmThisForm, grdThisGrid, gcmText
    ElseIf KeyCode = vbKeyA And (Shift And vbCtrlMask) Then
        ' Ctrl+A
        FlexGridSelectEntireGrid grdThisGrid, lngMaxColumnToSelect
    End If
End Sub

Private Sub FlexGridRemoveHighlighting(grdThisGrid As MSFlexGrid, lngStartCol As Long, lngEndCol As Long)
    Dim lngRow As Long, lngCol As Long
    
    With grdThisGrid
        For lngRow = 1 To .Rows - 1
            .Row = lngRow
            For lngCol = lngStartCol To lngEndCol
                .Col = lngCol
                .CellBackColor = 0
            Next lngCol
        Next lngRow
    End With
End Sub

Private Sub FlexGridSelectEntireGrid(grdThisGrid As MSFlexGrid, Optional lngMaxColumnToSelect As Long = -1)
    ' Select entire grid
    With grdThisGrid
        If .Rows > 1 Then
            .Row = 1
            .Col = 0
            .RowSel = .Rows - 1
            If lngMaxColumnToSelect = -1 Then lngMaxColumnToSelect = .Cols - 1
            
            If lngMaxColumnToSelect >= 0 And lngMaxColumnToSelect < .Cols Then
                .ColSel = lngMaxColumnToSelect
            Else
                If .Cols > 1 Then
                    .ColSel = 1
                Else
                    .ColSel = 0
                End If
            End If
        End If
    End With

End Sub

Private Function Get3LetterCodeState() As Boolean
    ' Returns true if the user wants 3 letter sequence residue symbols
    
    If cboNotation.ListIndex = 0 Then
        Get3LetterCodeState = False
    Else
        Get3LetterCodeState = True
    End If
    
End Function

Private Function GetCTerminusState() As ctgCTerminusGroupConstants
    Dim lngListIndex As Long
    
    lngListIndex = cboCTerminus.ListIndex
    
    If lngListIndex >= ctgHydroxyl And lngListIndex <= ctgNone Then
        GetCTerminusState = lngListIndex
    Else
        GetCTerminusState = ctgNone
    End If

End Function

Public Function GetIonMatchList(ByRef ThisIonMatchList() As Double, ByRef strIonMatchListCaption As String) As Long
    ' Returns the IonMatchList() since it is a private variable
    
    Dim lngRow As Long, lngCol As Long
    
    ReDim ThisIonMatchList(IonMatchListCount, 3)
    
    ' IonMatchList() is 1 based in the first dimension and 0-based in the second, using columns 0, 1, and 2
    For lngRow = 1 To IonMatchListCount
        For lngCol = 0 To 2
            ThisIonMatchList(lngRow, lngCol) = IonMatchList(lngRow, lngCol)
        Next lngCol
    Next lngRow
    
    strIonMatchListCaption = mIonMatchListDetails.Caption
    
    GetIonMatchList = IonMatchListCount
End Function


Private Function GetNTerminusState() As ntgNTerminusGroupConstants
    Dim lngListIndex As Long
    
    lngListIndex = cboNTerminus.ListIndex
    
    If lngListIndex >= ntgHydrogen And lngListIndex <= ntgNone Then
        GetNTerminusState = lngListIndex
    Else
        GetNTerminusState = ctgNone
    End If

End Function

Private Sub IonMatchListClear()
    InitializeIonMatchList 0
    InitializeIonListGrid
End Sub

Public Sub LoadCWSpectrumOptions(strSequenceFilePath As String)
    objSpectrum.LoadDataFromDisk strSequenceFilePath, False, True
End Sub

Public Function LoadIonListToMatch(Optional SeqFileNum As Integer = 0, Optional IonMatchListCountInFile As Long) As Boolean
    ' Loads Fragmentation Modelling sequence info from a file
    
    ' If SeqFileNum is not given (and thus 0), the user is prompted for the file from which to read the ion list
    ' Otherwise, a .Seq file is already open and the ion list in the file needs to be read
    '
    ' Returns true if a file is selected and one or more ions are read
    
    ' Can also read data from a .Dta file, in which case the parent ion will be properly recorded
    ' Finally, if the user chooses a _Dta.Txt file, then frmDtaTxtFileBrowser is called
    
    Static intFilterIndexSaved As Integer
    Dim strIonListFilename As String, strMessage As String
    Dim strLineIn As String, lngValuesToPopulate As Long
    Dim blnAllowCommaDelimeter As Boolean
    Dim IonListFileNum As Integer, lngPointsLoaded As Long
    Dim dblMinMass As Double, dblMaxMass As Double
    Dim lngCharLoc As Long
    Dim blnLoadingDTAFile As Boolean
    Dim blnSkippedParentIon As Boolean
    Dim strFilter As String
    
    If SeqFileNum = 0 Then
        ' Display the File Open dialog box.
        
        ' 1500 = All Files
        ' 1540 = Ion List Files, 1545 = .txt
        strFilter = LookupMessage(1540) & " (*." & LookupMessage(1545) & ")|*." & LookupMessage(1545) & "|DTA Files (*.dta)|*.dta|" & "Concatenated DTA files (*dta.txt)|*dta.txt|" & LookupMessage(1500) & " (*.*)|*.*"
        If intFilterIndexSaved = 0 Then intFilterIndexSaved = 1
        strIonListFilename = SelectFile(frmFragmentationModelling.hwnd, "Select File", gLastFileOpenSaveFolder, False, "", strFilter, intFilterIndexSaved)
        If Len(strIonListFilename) = 0 Then
            ' No file selected (or other error)
            LoadIonListToMatch = True
            Exit Function
        End If
    Else
        If IonMatchListCountInFile < 1 Then
            LoadIonListToMatch = False
            Exit Function
        End If
    End If
    
    If LCase(Right(strIonListFilename, 8)) = "_dta.txt" Then
        ' User chose a concatenated _dta.txt file, call frmDtaTxtFileBrowser instead
        ' This filetype is specific to PNNL and the Smith group
        frmDtaTxtFileBrowser.ReadDtaTxtFile strIonListFilename
        SetDtaTxtFileBrowserMenuVisibility frmDtaTxtFileBrowser.GetDataInitializedState()
        LoadIonListToMatch = True
        Exit Function
    End If
    
    On Error GoTo LoadIonListProblem
    
    ' Determine if commas are used for the decimal point in this locale
    If DetermineDecimalPoint() = "," Then
        blnAllowCommaDelimeter = False
    Else
        blnAllowCommaDelimeter = True
    End If
    
    If SeqFileNum = 0 Then
        ' Length of the progress bar is an estimate
        frmProgress.InitializeForm LookupMessage(930), 0, 4
        frmProgress.ToggleAlwaysOnTop True
    
        ' Open the file for input
        IonListFileNum = FreeFile()
        Open strIonListFilename For Input As #IonListFileNum
    
        ' First Determine number of data points in file
        ' Necessary since the IonMatchList() array is multi-dimensional and cannot be redimensioned without erasing old values
        lngValuesToPopulate = 0
        Do While Not EOF(IonListFileNum)
            Line Input #IonListFileNum, strLineIn
            
            If Len(strLineIn) > 0 Then
                lngValuesToPopulate = lngValuesToPopulate + 1
            End If
        Loop
        Close #IonListFileNum
        
        frmProgress.InitializeForm LookupMessage(930), 0, lngValuesToPopulate
    Else
        lngValuesToPopulate = IonMatchListCountInFile
    End If
    
    ' Initialize the IonMatchList() Array
    InitializeIonMatchList lngValuesToPopulate
    lngPointsLoaded = 0
    
    dblMaxMass = LowestValueForDoubleDataType
    dblMinMass = HighestValueForDoubleDataType
    
    If SeqFileNum = 0 Then
        ' Now re-open the file and import the data
        IonListFileNum = FreeFile
        Open strIonListFilename For Input As #IonListFileNum
    
        If LCase(Right(strIonListFilename, 4)) = ".dta" Then
            ' User chose a .Dta file
            ' Need to set the following to True
            blnLoadingDTAFile = True
        End If
    Else
        IonListFileNum = SeqFileNum
    End If
    
    Do While Not EOF(IonListFileNum)
        If lngPointsLoaded Mod 250 = 0 And SeqFileNum = 0 Then
            frmProgress.UpdateProgressBar lngPointsLoaded
            If KeyPressAbortProcess > 1 Then Exit Do
        End If
        
        Line Input #IonListFileNum, strLineIn
        strLineIn = Trim(strLineIn)
        
        If blnLoadingDTAFile And Not blnSkippedParentIon Then
            If Len(strLineIn) > 0 Then
                If IsNumeric(Left(strLineIn, 1)) Then
                    ' Skip the first set of data since it is the parent ion (MH) and charge
                    lngCharLoc = InStr(strLineIn, " ")
                    If lngCharLoc > 0 Then
                        strLineIn = Left(strLineIn, lngCharLoc - 1)
                        txtPrecursorIonMass = strLineIn
                    End If
                    blnSkippedParentIon = True
                    strLineIn = ""
                End If
            End If
        End If
        
        If Len(strLineIn) > 0 Then
            If UCase(Left(strLineIn, 4)) = "FRAG" Then Exit Do
            AddIonPairToIonList strLineIn, lngPointsLoaded, blnAllowCommaDelimeter, dblMinMass, dblMaxMass
        End If
    Loop
    
    If SeqFileNum = 0 Then Close #IonListFileNum
    
    If KeyPressAbortProcess = 0 And lngPointsLoaded > 0 Then
        LoadIonListToMatch = True
        
        If Abs((dblMaxMass - dblMinMass) / lngPointsLoaded) < ION_SEPARATION_TOLERANCE_TO_AUTO_ENABLE_BINNING Then
            ' Average spacing between ions is less than 0.2 Da, turn on binning
            frmIonMatchOptions.chkGroupSimilarIons.value = vbChecked
            frmIonMatchOptions.txtGroupIonMassWindow.Text = "0.2"
        End If

        ' Reset the alignment offset
        txtAlignment = "0"
    Else
        If SeqFileNum = 0 And lngPointsLoaded > 0 Then
            MsgBox LookupMessage(940), vbInformation + vbOKOnly, LookupMessage(945)
        End If
        LoadIonListToMatch = False
    End If
    
    If SeqFileNum = 0 Then frmProgress.HideForm
        
    Exit Function
    
LoadIonListProblem:
    If SeqFileNum = 0 Then Close
    strMessage = LookupMessage(900) & ": " & strIonListFilename
    strMessage = strMessage & vbCrLf & Err.Description
    MsgBox strMessage, vbOKOnly + vbExclamation, LookupMessage(350)

End Function

Private Function MassToBinNumber(ThisMass As Double, StartMass As Double, MassResolution As Double) As Long
    Dim WorkingMass As Double
    
    ' First subtract StartMass from ThisMass
    ' For example, if StartMass is 500 and ThisMass is 500.28, then WorkingMass = 0.28
    ' Or, if StartMass is 500 and ThisMass is 530.83, then WorkingMass = 30.83
    WorkingMass = ThisMass - StartMass
    
    ' Now, dividing WorkingMass by MassResolution and rounding to nearest integer
    '  actually gives the bin
    ' For example, given WorkingMass = 0.28 and MassResolution = 0.1, Bin = CLng(2.8) + 1 = 4
    ' Or, given WorkingMass = 30.83 and MassResolution = 0.1, Bin = CLng(308.3) + 1 = 309
    MassToBinNumber = CLng(WorkingMass / MassResolution) + 1
        
End Function
    
Private Sub MatchIons()
    ' For each of the ions in FragSpectrumDetails(), look in IonMatchList() to see
    '  if any of the ions is within tolerance
    ' If it is, change the background color of the corresponding cell in grdFragMasses accordingly

    Dim lngIonIndex As Long, lngResidueIndex As Long
    Dim lngRowIndex As Long
    Dim eIonType As itIonTypeConstants
    Dim intChargeIndex As Integer
    Dim lngCellBackColor As Long
    Dim dblMatchTolerance As Double
    Dim dblFragmentMass As Double
    Dim lngIndexInIonListGrid As Long
    
    Dim udtResidueMatched() As udtResidueMatchedType
    Dim lngResidueCount As Long
    
    ' Variables for computing match score
    Dim dblScore As Double
    Dim lngMatchCount As Long           ' Number of ions in user data matching a predicted ion
    Dim dblIntensitySum As Double       ' Sum of the intensities for all matching ions (normalized intensity in user data)
    Dim dblBeta As Double               ' Incremented by 0.075 for each successive b ion and also each successive y ion
    Dim dblRho As Double                ' Adjustment for presence of standard immonium ions:
                                        '   Looks for 110, 120, and 136 m/z in user data (His, Phe, and Tyr immonium ions)
                                        '   If m/z is present, examines sequence being matched.  If corresponding amino acid is present,
                                        '   increments rho by 0.15.  If amino acid is absent, decrements rho by 0.15.  No change if m/z is not found in user data
                                        
    
    If grdIonList.Rows <= 1 Or FragSpectrumIonCount < 1 Then Exit Sub
    
    If Not mFragMatchSettingsChanged Then Exit Sub
    
    mFragMatchSettingsChanged = False
    
    If Not (grdIonList.Visible Or mAutoAligning) Then Exit Sub
    
    dblMatchTolerance = Val(txtIonMatchingWindow)
    If dblMatchTolerance < 0 Or dblMatchTolerance > 100 Then
        dblMatchTolerance = 0.5
    End If
    
    If Not mAutoAligning Then Me.MousePointer = vbHourglass
    
    ' Initialize variables used to compute score
    lngMatchCount = 0
    dblIntensitySum = 0
    dblBeta = 0
    dblRho = 0
    
    ' Initialize udtResidueMatched()
    lngResidueCount = objMwtWin.Peptide.GetResidueCount
    ReDim udtResidueMatched(lngResidueCount)
    
    ' Hide grdFragMasses and grdIonList to prevent screen updates
    '  from slowing down the computer during matching
    grdFragMasses.Visible = False
    grdIonList.Visible = False
     
    ' Remove all background coloring from grdIonList
    FlexGridRemoveHighlighting grdIonList, 0, ION_LIST_COL_COUNT - 1
    
    ' Clear the ilgSymbolMatch column in grdIonList
    With grdIonList
        For lngRowIndex = 1 To .Rows - 1
            .TextMatrix(lngRowIndex, ilgSymbolMatch) = ""
        Next lngRowIndex
    End With
    
    lngMatchCount = 0
    For lngIonIndex = 0 To FragSpectrumIonCount - 1
        With FragSpectrumDetails(lngIonIndex)
            If .Mass > 0 Then
                lngIndexInIonListGrid = SearchForIonInIonListGrid(.Mass, dblMatchTolerance)
                
                If lngIndexInIonListGrid >= 0 Then
                    lngMatchCount = lngMatchCount + 1
                    
                    grdFragMasses.Row = FragSpectrumGridLocs(lngIonIndex).Row
                    grdFragMasses.Col = FragSpectrumGridLocs(lngIonIndex).Col
                    If .IsShoulderIon Then
                        lngCellBackColor = &H80FFFF      ' Light Yellow
                        
                        ' Only change to yellow if currently white
                        If grdFragMasses.CellBackColor = 0 Then
                            grdFragMasses.CellBackColor = lngCellBackColor
                        End If
                    Else
                        lngCellBackColor = vbCyan
                        grdFragMasses.CellBackColor = lngCellBackColor
                    End If
                    
                    dblIntensitySum = dblIntensitySum + Abs(.Intensity)
                    
                    ' Display the symbol of the match in the ilgSymbolMatch column
                    grdIonList.TextMatrix(lngIndexInIonListGrid, ilgSymbolMatch) = .Symbol
                    
                    ' Also, color the corresponding row in grdIonList
                    ' Fill in order of 2, 1, 0 so cursor ends up in 0'th column
                    grdIonList.Row = lngIndexInIonListGrid
                    grdIonList.Col = ilgSymbolMatch: grdIonList.CellBackColor = lngCellBackColor
                    grdIonList.Col = ilgIntensity: grdIonList.CellBackColor = lngCellBackColor
                    grdIonList.Col = ilgMass: grdIonList.CellBackColor = lngCellBackColor
                    
                    ' Lastly, update udtResidueMatched()
                    udtResidueMatched(.SourceResidueNumber).IonHit(.IonType, .Charge) = True
                End If

            End If
        End With
    Next lngIonIndex
            
    ' Now examine udtResidueMatched() to determine dblBeta
    ' Beta is incremented by 0.075 for each successive b, b++, y, y++, c, c++, z, or z++ ion
    For eIonType = itBIon To itZIon
        For intChargeIndex = 1 To 2
            For lngResidueIndex = 0 To lngResidueCount - 2
                If udtResidueMatched(lngResidueIndex).IonHit(eIonType, intChargeIndex) Then
                    If udtResidueMatched(lngResidueIndex + 1).IonHit(eIonType, intChargeIndex) Then
                        dblBeta = dblBeta + 0.075
                    End If
                End If
            Next lngResidueIndex
        Next intChargeIndex
    Next eIonType
    
    ' Finally, examine grdIonList to look for important immonium ions, adjusting rho as needed
    
    ' Look for Histidine (110)
    dblFragmentMass = dblHistidineFW - dblImmoniumMassDifference
    If Not WithinTolerance(110, dblFragmentMass, 2) Then dblFragmentMass = 110
    MatchIonAdjustRho dblRho, dblFragmentMass, dblMatchTolerance, "His"
    
    ' Look for Phenylalanine (120)
    dblFragmentMass = dblPhenylalanineFW - dblImmoniumMassDifference
    If Not WithinTolerance(120, dblFragmentMass, 2) Then dblFragmentMass = 120
    MatchIonAdjustRho dblRho, dblFragmentMass, dblMatchTolerance, "Phe"
    
    ' Look for Tyrosine (136)
    dblFragmentMass = dblTyrosineFW - dblImmoniumMassDifference
    If Not WithinTolerance(136, dblFragmentMass, 2) Then dblFragmentMass = 136
    MatchIonAdjustRho dblRho, dblFragmentMass, dblMatchTolerance, "Tyr"
    
    ' Compute the score
    If FragSpectrumIonCount > 0 Then
        ' Compute Match Score
        ' ToDo: whether 1 - dblRho or 1 + dblRho is correct
        dblScore = dblIntensitySum * lngMatchCount * (1 + dblBeta) * (1 + dblRho) / FragSpectrumIonCount
    Else
        dblScore = 0
    End If

    If Not mAutoAligning Then
        grdFragMasses.Visible = True
        grdIonList.Visible = True
    
        UpdateMatchCountAndScoreWork lngMatchCount, dblScore
    Else
        lngAlignmentMatchCount = lngMatchCount
        dblAlignmentScore = dblScore
    End If

    If Not mAutoAligning Then Me.MousePointer = vbDefault
    
End Sub

Private Function MatchIonAdjustRho(ByRef dblRho As Double, dblFragmentMass As Double, dblMatchTolerance As Double, strResidue3LetterSymbol As String) As Long
    ' Returns -1 if no match
    ' Otherwise, returns the row index of the match in grdIonList
    
    Dim lngIndexInIonListGrid As Long
    Dim lngResidueIndex As Long
    
    lngIndexInIonListGrid = SearchForIonInIonListGrid(dblFragmentMass, dblMatchTolerance)
    
    ' Only adjust rho if intensity of ion in grdIonList is greater than 40% of the
    ' maximum intensity used for normalizing
    
    If lngIndexInIonListGrid >= 0 Then
        
        If Val(grdIonList.TextMatrix(lngIndexInIonListGrid, ilgIntensity)) >= 0.4 * Val(frmIonMatchOptions.txtNormalizedIntensity) Then
            
            ' See if residue exists in target sequence
            If objMwtWin.Peptide.GetResidueCountSpecificResidue(strResidue3LetterSymbol, True) > 0 Then
                ' Immonium mass found and residue is present in sequence
                ' Increase Rho by 0.15
                dblRho = dblRho + 0.15
                
                For lngResidueIndex = 1 To objMwtWin.Peptide.GetResidueCount
                    If objMwtWin.Peptide.GetResidueSymbolOnly(lngResidueIndex) = strResidue3LetterSymbol Then
                        ' Color given cell in grdFragMasses
                        With grdFragMasses
                            .Row = lngResidueIndex
                            .Col = 1
                            .CellBackColor = vbCyan
                        End With
                    End If
                Next lngResidueIndex
                
                ' Also color cell in grdIonList
                With grdIonList
                    .Row = lngIndexInIonListGrid
                    .Col = ilgSymbolMatch
                    .CellBackColor = vbCyan
                    .Col = ilgIntensity
                    .CellBackColor = vbCyan
                    .Col = ilgMass
                    .CellBackColor = vbCyan
                    
                End With
            Else
                ' Immonium mass found, but residue was not present in sequence
                ' Decrease Rho by 0.15
                dblRho = dblRho - 0.15
                
                ' Color cell in grdIonList red
                With grdIonList
                    .Row = lngIndexInIonListGrid
                    .Col = ilgSymbolMatch
                    .CellBackColor = &HC0C0FF       ' Light Red
                    .Col = ilgIntensity
                    .CellBackColor = &HC0C0FF       ' Light Red
                    .Col = ilgMass
                    .CellBackColor = &HC0C0FF       ' Light Red
                End With
                
                ' Set lngIndexInIonListGrid to -1
                lngIndexInIonListGrid = -1
            End If
        Else
            lngIndexInIonListGrid = -1
        End If
    End If
End Function

Private Sub NormalizeIonMatchListWrapper()
    Dim dblOriginalMaximumIntensity As Double
    
    If BinnedDataCount > 0 Then
        NormalizeIonMatchList BinnedData(), BinnedDataCount, dblOriginalMaximumIntensity
    Else
        NormalizeIonMatchList IonMatchList(), IonMatchListCount, dblOriginalMaximumIntensity
        mIonMatchListDetails.OriginalMaximumIntensity = dblOriginalMaximumIntensity
    End If
End Sub

Private Sub NormalizeIonMatchList(ThisIonList() As Double, ThisIonListCount As Long, ByRef dblOriginalMaximumIntensity As Double)
    Dim lngIndex As Long, lngRegionIndex As Long
    Dim lngStopIndex As Long
    Dim lngIonPointerArray() As Long, lngPointerArrayCount As Long
    Dim dblMassRegionIndices() As Double      ' 2D Array (1-based in the 1st dimension and 0-based in the second)
                                              ' First dimension is mass value of window edge, 2nd dimension is index in ThisIonList() that starts this window
    Dim intMassRegionCount As Integer, intMassRegionTrack As Integer
    Dim lngTopMostValuesToUse As Long
    Dim dblMassValue As Double
    Dim dblMassMinimum As Double, dblMassMaximum As Double
    Dim dblMassRegionWidth As Double
    Dim dblMaximumIntensity As Double
    Dim dblNormalizedIntensity As Double
    Dim boolRemovePrecursorIon As Boolean
    Dim dblPrecursorIonMass As Double, dblPrecursorIonMassWindow As Double
    Dim boolPrecursorRemoved As Boolean
    Dim blnShowProgress As Boolean
    
    If ThisIonListCount > 3000 Then blnShowProgress = True
    dblOriginalMaximumIntensity = 0
    If ThisIonListCount = 0 Then Exit Sub
    
    If blnShowProgress Then
        frmProgress.InitializeForm LookupMessage(950), 0, 4
        frmProgress.ToggleAlwaysOnTop True
    End If
    
    ' Fill tolerance variables
    If IsNumeric(txtPrecursorIonMass) Then
        dblPrecursorIonMass = Val(txtPrecursorIonMass)
    Else
        dblPrecursorIonMass = 0
    End If
    
    boolRemovePrecursorIon = cChkBox(chkRemovePrecursorIon)
    dblPrecursorIonMassWindow = txtPrecursorMassWindow
    If dblPrecursorIonMassWindow < 0 Or dblPrecursorIonMassWindow > 100 Then
        dblPrecursorIonMassWindow = 2
    End If
    
    With frmIonMatchOptions
        
        intMassRegionCount = .txtMassRegions
        If intMassRegionCount < 1 Or intMassRegionCount > 1000 Then
            intMassRegionCount = 1
        End If
        
        lngTopMostValuesToUse = .txtIonCountToUse
        If lngTopMostValuesToUse < 10 Or lngTopMostValuesToUse > 32000 Then
            lngTopMostValuesToUse = 200
        End If
        
        dblNormalizedIntensity = .txtNormalizedIntensity
        If dblNormalizedIntensity < 1 Or dblNormalizedIntensity > 32000 Then
            dblNormalizedIntensity = 100
        End If
    End With
    
    ' Initialize the Pointer Array
    lngPointerArrayCount = ThisIonListCount
    ReDim lngIonPointerArray(lngPointerArrayCount)
    For lngIndex = 1 To lngPointerArrayCount
        lngIonPointerArray(lngIndex) = lngIndex
    Next lngIndex
    
    ' Find the Mass Limits
    FindMassLimits ThisIonList(), ThisIonListCount, dblMassMinimum, dblMassMaximum
    
    ' Initialize dblMassRegionIndices()
    ReDim dblMassRegionIndices(intMassRegionCount + 1, 2)
    dblMassRegionWidth = (dblMassMaximum - dblMassMinimum) / intMassRegionCount
    
    For lngIndex = 1 To intMassRegionCount
        dblMassRegionIndices(lngIndex, 0) = dblMassMinimum + (lngIndex - 1) * dblMassRegionWidth
    Next lngIndex
    
    dblMassRegionIndices(1, 1) = 1
    intMassRegionTrack = 2
    dblMassRegionIndices(intMassRegionCount + 1, 0) = ThisIonList(ThisIonListCount, 0)
    dblMassRegionIndices(intMassRegionCount + 1, 1) = ThisIonListCount
    
    ' Initialize the 3rd column of ThisIonList() by setting to the actual intensity
    '  unless within tolerance from the PrecursorIonMass, in which case the intensity is set to -1
    ' At the same time, determine the index of the start of each mass region
    dblMassMinimum = HighestValueForDoubleDataType
    dblMassMaximum = LowestValueForDoubleDataType
    For lngIndex = 1 To ThisIonListCount
        dblMassValue = ThisIonList(lngIndex, ilaMass)
        ThisIonList(lngIndex, ilaNormalizedIntensity) = ThisIonList(lngIndex, ilaIntensity)
        If dblPrecursorIonMass > 0 And boolRemovePrecursorIon Then
            If WithinTolerance(dblMassValue, dblPrecursorIonMass, dblPrecursorIonMassWindow) Then
                ThisIonList(lngIndex, ilaNormalizedIntensity) = -1
                boolPrecursorRemoved = True
            End If
        End If
        
        If intMassRegionTrack <= intMassRegionCount Then
            If dblMassValue >= dblMassRegionIndices(intMassRegionTrack, 0) Then
                dblMassRegionIndices(intMassRegionTrack, 1) = lngIndex
                intMassRegionTrack = intMassRegionTrack + 1
            End If
        End If
    Next lngIndex
    
    ' Update lblPrecursorStatus
    If boolPrecursorRemoved Then
        lblPrecursorStatus.Caption = LookupLanguageCaption(12390, "Precursor removed") & ": " & Trim(CStr(Round(dblPrecursorIonMass, 0))) & "±" & Format(dblPrecursorIonMassWindow, "0.0")
    Else
        If boolRemovePrecursorIon Then
            lblPrecursorStatus.Caption = LookupLanguageCaption(12385, "Precursor not found")
        Else
            lblPrecursorStatus.Caption = LookupLanguageCaption(12395, "Precursor not removed")
        End If
    End If
    
    lngPointerArrayCount = ThisIonListCount
    
    ' Next, sort the ions by intensity (descending)
    ShellSortIonList ThisIonList(), lngIonPointerArray(), 1, lngPointerArrayCount, ilaNormalizedIntensity, blnShowProgress
    
    ' Next, set the intensities of ions below the top-most ion count to -1 (so that they'll be ignored
    ' If lngTopMostValuesToUse > lngPointerArrayCount then nothing is set to -1
    For lngIndex = lngTopMostValuesToUse + 1 To lngPointerArrayCount
        ThisIonList(lngIonPointerArray(lngIndex), ilaNormalizedIntensity) = -1
    Next lngIndex
    
    If blnShowProgress Then
        frmProgress.InitializeForm LookupMessage(950), 0, 4
        frmProgress.UpdateProgressBar 2
        frmProgress.UpdateCurrentSubTask LookupMessage(960)
    End If
    
    ' Next, normalize the intensities in each mass region
    ' For each region, find the maximum intensity
    ' Then normalize intensities in given region to maximum intensity
    For lngRegionIndex = 1 To intMassRegionCount
        dblMaximumIntensity = LowestValueForDoubleDataType
        For lngIndex = dblMassRegionIndices(lngRegionIndex, 1) To dblMassRegionIndices(lngRegionIndex + 1, 1)
            If ThisIonList(lngIndex, ilaNormalizedIntensity) > dblMaximumIntensity Then dblMaximumIntensity = ThisIonList(lngIndex, ilaNormalizedIntensity)
        Next lngIndex
        If dblMaximumIntensity <= 0 Then
            ' All ions in this region were below tolerance
            dblMaximumIntensity = 1
        Else
            If dblMaximumIntensity > dblOriginalMaximumIntensity Then
                dblOriginalMaximumIntensity = dblMaximumIntensity
            End If
        End If
        ' Finally, step through list, normalizing intensities if > 0
        lngStopIndex = dblMassRegionIndices(lngRegionIndex + 1, 1) - 1
        If lngRegionIndex = intMassRegionCount Then
            lngStopIndex = lngStopIndex + 1
        End If
        For lngIndex = dblMassRegionIndices(lngRegionIndex, 1) To lngStopIndex
            If ThisIonList(lngIndex, ilaNormalizedIntensity) > 0 Then
                ' Round Intensity to 3 decimal places
                ThisIonList(lngIndex, ilaNormalizedIntensity) = Round(ThisIonList(lngIndex, ilaNormalizedIntensity) / dblMaximumIntensity * dblNormalizedIntensity, 3)
            End If
        Next lngIndex
    Next lngRegionIndex

    FillIonMatchGridWrapper
    
    If blnShowProgress Then
        frmProgress.InitializeForm LookupMessage(950), 0, 4
        frmProgress.UpdateProgressBar 3
        frmProgress.UpdateCurrentSubTask LookupMessage(970)
    End If
    
    UpdateMassesGridAndSpectrumWrapper
    
    If KeyPressAbortProcess > 1 Then
        MsgBox LookupMessage(940), vbInformation + vbOKOnly, LookupMessage(945)
    End If
    
    frmProgress.HideForm
End Sub

Private Sub PasteIonMatchList()
    Dim strIonList As String, strLineWork As String
    Dim lngPointsLoaded As Long, blnAllowCommaDelimeter As Boolean
    Dim lngCrLfLoc As Long, lngValuesToPopulate As Long
    Dim lngIndex As Long, lngStartIndex As Long
    Dim dblMinMass As Double, dblMaxMass As Double
    Dim strDelimeter As String, intDelimeterLength As Integer
    
On Error GoTo PasteIonMatchListErrorHandler

    ' Warn user about replacing data
    If Not QueryExistingIonsInList() Then Exit Sub
    
    ' Grab text from clipboard
    strIonList = GetClipboardTextSmart()
    
    If Len(strIonList) = 0 Then
        MsgBox LookupMessage(980), vbInformation + vbOKOnly, LookupMessage(985)
        Exit Sub
    End If
        
    ' Determine if commas are used for the decimal point in this locale
    If DetermineDecimalPoint() = "," Then
        blnAllowCommaDelimeter = False
    Else
        blnAllowCommaDelimeter = True
    End If
        
    frmProgress.InitializeForm LookupMessage(990), 0, Len(strIonList)
    frmProgress.ToggleAlwaysOnTop True
    frmProgress.UpdateCurrentSubTask LookupMessage(1000)
    lngPointsLoaded = 0
    
    ' First Determine number of data points in strIonList
    ' Necessary since the IonMatchList() array is multi-dimensional and cannot be redimensioned without erasing old values
    
    ' First look for the first occurrence of a valid delimeter
    lngCrLfLoc = ParseStringFindCrlfIndex(strIonList, intDelimeterLength)
    
    If lngCrLfLoc > 0 Then
        ' Record the first character of the delimeter (if vbCrLf, then recording vbCr), otherwise, recording just vbCr or just vbLf)
        strDelimeter = Mid(strIonList, lngCrLfLoc, 1)
        
        ' Now determine the number of occurrences of the delimeter
        lngValuesToPopulate = 0
        For lngIndex = 1 To Len(strIonList)
            If lngIndex Mod 500 = 0 Then
                frmProgress.UpdateProgressBar lngIndex
                If KeyPressAbortProcess > 1 Then Exit For
            End If
            
            If Mid(strIonList, lngIndex, 1) = strDelimeter Then
                lngValuesToPopulate = lngValuesToPopulate + 1
            End If
        Next lngIndex
        
        If lngValuesToPopulate > 1 And KeyPressAbortProcess = 0 Then
            ' Update frmProgress to use the correct progress bar size
            frmProgress.InitializeForm LookupMessage(990), 0, lngValuesToPopulate
            frmProgress.UpdateCurrentSubTask LookupMessage(1010)
        End If
        
        ' Initialize IonMatchList
        InitializeIonMatchList lngValuesToPopulate
        lngPointsLoaded = 0
        
        dblMaxMass = LowestValueForDoubleDataType
        dblMinMass = HighestValueForDoubleDataType
        
        ' Actually parse the data
        ' I process the list using the Mid() function since this is the fastest method
        ' However, Mid() becomes slower when the index value reaches 1000 or more (roughly)
        '  so I discard already-parsed data from strIonList every 1000 characters (approximately)
        Do While Len(strIonList) > 0
            lngStartIndex = 1
            For lngIndex = 1 To Len(strIonList)
                If lngIndex Mod 100 = 0 Then
                    frmProgress.UpdateProgressBar lngPointsLoaded
                    If KeyPressAbortProcess > 1 Then Exit For
                End If
                
                If Mid(strIonList, lngIndex, 1) = strDelimeter Then
                    strLineWork = Mid(strIonList, lngStartIndex, lngIndex - 1)
        
                    If Len(strLineWork) > 0 Then
                        AddIonPairToIonList strLineWork, lngPointsLoaded, blnAllowCommaDelimeter, dblMinMass, dblMaxMass
                    End If
                    
                    lngStartIndex = lngIndex + intDelimeterLength
                    lngIndex = lngIndex + intDelimeterLength - 1
                    
                    If lngIndex > 1000 Then
                        ' Reduce the size of strIonList since the Mid() function gets slower with longer strings
                        strIonList = Mid(strIonList, lngIndex + 1)
                        lngIndex = 1
                        Exit For
                    End If
                End If
            Next lngIndex
            If lngIndex > Len(strIonList) Then Exit Do
        Loop
    End If
    
    If lngPointsLoaded > 0 And KeyPressAbortProcess = 0 Then
        If Abs((dblMaxMass - dblMinMass) / lngPointsLoaded) < ION_SEPARATION_TOLERANCE_TO_AUTO_ENABLE_BINNING Then
            ' Average spacing between ions is less than 0.2 Da, turn on binning
            frmIonMatchOptions.chkGroupSimilarIons.value = vbChecked
        End If
        
        ' Reset the alignment offset
        txtAlignment = "0"
            
        UpdateIonMatchListWrapper
        frmProgress.HideForm
    Else
        frmProgress.HideForm
        If KeyPressAbortProcess = 0 Then
            MsgBox LookupMessage(1020), vbInformation + vbOKOnly, LookupMessage(985)
        Else
            MsgBox LookupMessage(940), vbInformation + vbOKOnly, LookupMessage(945)
        End If
    End If
    
    Exit Sub
    
PasteIonMatchListErrorHandler:
    frmProgress.HideForm
    GeneralErrorHandler "frmFragmentationModelling|ParseIonMatchList", Err.Number, Err.Description

End Sub

Public Sub PasteNewSequence(ByVal strNewSequence As String, bln3LetterCode As Boolean)
    
    ' Validate strNewSequence
    objMwtWin.Peptide.SetSequence strNewSequence, , , bln3LetterCode

    txtSequence = objMwtWin.Peptide.GetSequence(bln3LetterCode)
    
End Sub

Private Sub PeptideSequenceKeyPressHandler(txtThisTextBox As TextBox, KeyAscii As Integer)
    ' Checks KeyAscii to see if it's valid
    ' If it isn't, it is set to 0
    
    If Not objMwtWin.IsModSymbol(Chr(KeyAscii)) Then
        Select Case KeyAscii
        Case 1          ' Ctrl+A -- select entire textbox
            txtThisTextBox.SelStart = 0
            txtThisTextBox.SelLength = Len(txtThisTextBox.Text)
            KeyAscii = 0
        Case 24, 3, 22  ' Cut, Copy, Paste is allowed
        Case 26
            ' Ctrl+Z = Undo
            KeyAscii = 0
            txtThisTextBox.Text = GetMostRecentTextBoxValue()
        Case 8          ' Backspace key is allowed
        Case 48 To 57   ' Numbers are not allowed       ' ToDo: Possibly allow them for side-chains
            KeyAscii = 0
        Case 32         ' Spaces are allowed
        Case 40 To 41   ' Parentheses are not allowed       ' ToDo: Possibly allow them for side-chains
            KeyAscii = 0
        Case 43:        ' Plus sign is not allowed
            KeyAscii = 0
        Case 45:        ' Negative sign is allowed
        Case 44, 46:    ' Decimal point (. or ,) is allowed
        Case 65 To 90, 97 To 122    ' Characters are allowed
        Case 95:        ' Underscore is not allowed
            KeyAscii = 0
        Case Else
            KeyAscii = 0
        End Select
    End If
    
End Sub

Private Sub PopulateComboBoxes()

    Dim intIndex As Integer
    
    On Error GoTo PopulateComboBoxesErrorHandler
    
    PopulateComboBox cboNotation, True, "1 letter notation|3 letter notation", 1
    
    cboDoubleCharge.Clear
    With cboDoubleCharge
        For intIndex = 1 To 16
            .AddItem CIntSafeDbl((intIndex - 1) * 100)
        Next intIndex
        .ListIndex = 8
    End With
    
    cboTripleCharge.Clear
    With cboTripleCharge
        For intIndex = 1 To cboDoubleCharge.ListCount
            .AddItem cboDoubleCharge.List(intIndex - 1)
        Next intIndex
        .ListIndex = 9
    End With
    
    PopulateComboBox cboNTerminus, True, "H (hydrogen)|HH+ (protonated)|C2OH3 (acetyl)|C5O2NH6 (pyroglu)|CONH2 (carbamyl)|C7H6NS (PTC)|(none)", 0
    
    PopulateComboBox cboCTerminus, True, "OH (hydroxyl)|NH2 (amide)|(none)", 0
        
    With lstIonsToModify
        .Clear
        .AddItem UCase(LookupLanguageCaption(12600, "a"))
        .AddItem UCase(LookupLanguageCaption(12610, "b"))
        .AddItem UCase(LookupLanguageCaption(12620, "y"))
        .AddItem UCase(LookupLanguageCaption(12630, "c"))
        .AddItem UCase(LookupLanguageCaption(12640, "z"))
        .Selected(0) = False
        .Selected(1) = True
        .Selected(2) = True
    End With
    
    With cboMHAlternate
        .Clear
        .AddItem "[M+H]1+"
        For intIndex = 2 To 9
            .AddItem "[M+" & Trim(intIndex) & "H]" & Trim(intIndex) & "+"
        Next intIndex
        .ListIndex = 1    ' Charge of 2+
    End With
    Exit Sub
    
PopulateComboBoxesErrorHandler:
    GeneralErrorHandler "frmFragmentationModelling|PopulateComboBoxes", Err.Number, Err.Description

End Sub

Private Sub PositionControls()
    Dim lngTopAdjust As Long, blnIonMatchListShown As Boolean
    Dim lngPreferredTop As Long, lngMinimumTop As Long
    Dim lngPreferredHeight As Long
    Dim lngPreferredWidth As Long
    Dim blnSkipWidthAdjust As Boolean, blnSkipHeightAdjust As Boolean
    
    If Me.WindowState = vbMinimized Then Exit Sub
    
    blnIonMatchListShown = grdIonList.Visible
    If Me.Height < 6500 And Not blnIonMatchListShown Then
        lngTopAdjust = 0
    Else
        lngTopAdjust = 340
    End If
    
    If Me.Height < 2660 + lngTopAdjust Then
        'Me.Height = 2660 + lngTopAdjust
        ' Do not position any controls; form too small
        blnSkipHeightAdjust = True
    End If
    
    If blnIonMatchListShown Then
        lngPreferredWidth = 8000
    Else
        lngPreferredWidth = 6700
    End If
    
    If Me.Width < lngPreferredWidth Then
        'Me.Width = 6700
        ' Do not position any controls; form too small
        blnSkipWidthAdjust = True
    End If
    
    lblSequence.Top = 180
    lblSequence.Left = 60
    txtSequence.Top = 120
    txtSequence.Left = 1440
    If Not blnSkipWidthAdjust Then txtSequence.Width = Me.Width - txtSequence.Left - cmdMatchIons.Width - 340
    txtSequence.Height = 315 + lngTopAdjust
    
    If Not blnSkipWidthAdjust Then
        cmdMatchIons.Left = txtSequence.Left + txtSequence.Width + 120
    End If
    
    If Not blnSkipHeightAdjust Then
        cmdMatchIons.Top = txtSequence.Top + (txtSequence.Height / 2) - (cmdMatchIons.Height / 2)
    End If
    
    With fraMassInfo
        .Top = 460 + lngTopAdjust
        .Left = 2640
    End With
    lngTopAdjust = lngTopAdjust + fraMassInfo.Height
    
    cboNotation.Top = fraMassInfo.Top + 120
    cboNotation.Left = 120
    
    lblMH.Left = cboMHAlternate.Left + 60
    
    With grdIonList
        .Top = fraMassInfo.Top + fraMassInfo.Height + 100
        lngPreferredWidth = 2280
        If Me.Width > 10000 Then
            lngPreferredWidth = lngPreferredWidth + (Me.Width - 10000) / 4
            If lngPreferredWidth > 3270 Then lngPreferredWidth = 3270
        End If
        
        .Width = lngPreferredWidth
        If Not blnSkipWidthAdjust Then .Left = Me.Width - .Width - 225
        If Not blnSkipHeightAdjust Then .Height = Me.Height - .Top - 900
    End With
        
    fraTerminii.Top = grdIonList.Top
    fraTerminii.Left = 60
    
    fraIonTypes.Top = fraTerminii.Top + fraTerminii.Height + 80
    fraIonTypes.Left = fraTerminii.Left
    
    fraNeutralLosses.Top = fraIonTypes.Top + fraIonTypes.Height + 80
    fraNeutralLosses.Left = fraTerminii.Left
    
    fraCharge.Top = fraNeutralLosses.Top + fraNeutralLosses.Height + 80
    fraCharge.Left = fraTerminii.Left
    
    ' Preferred top for fraIonMatching
    lngPreferredTop = Me.Height - fraIonMatching.Height - 860
    lngMinimumTop = fraCharge.Top + fraCharge.Height + 80
    If lngPreferredTop < lngMinimumTop Then lngPreferredTop = lngMinimumTop
    
    With fraIonMatching
        If Not blnSkipHeightAdjust Then .Top = lngPreferredTop
        .Left = fraTerminii.Left
    End With
        
    With fraIonStats
        .Top = fraIonMatching.Top
        .Left = fraIonMatching.Left + fraIonMatching.Width + 90
        lngPreferredWidth = fraIonMatching.Width
        If .Left + lngPreferredWidth > grdIonList.Left - 80 Then
            lngPreferredWidth = grdIonList.Left - 80 - .Left
            If lngPreferredWidth < 1500 Then
                lngPreferredWidth = 1500
            End If
        End If
        .Width = lngPreferredWidth
        
        ' Remove the border (I keep it on in the design window so I can see where the labels are
        lblIonLoadedCount.BorderStyle = 0
        lblBinAndToleranceCounts.BorderStyle = 0
        lblPrecursorStatus.BorderStyle = 0
        lblMatchCount.BorderStyle = 0
        lblScore.BorderStyle = 0
        
        lblIonLoadedCount.Width = lngPreferredWidth - 250
        lblBinAndToleranceCounts.Width = lblIonLoadedCount.Width
        lblPrecursorStatus.Width = lblIonLoadedCount.Width
        lblMatchCount.Width = lblIonLoadedCount.Width
        lblScore.Width = lblIonLoadedCount.Width
    End With
    
    With grdFragMasses
        .Left = fraMassInfo.Left
        .Top = grdIonList.Top
        
        ' Preferred height for grdFragMasses
        lngPreferredHeight = grdIonList.Height
        If blnIonMatchListShown Then
            If .Top + lngPreferredHeight + 80 > lngPreferredTop Then
                lngPreferredHeight = fraIonMatching.Top - .Top - 80
            End If
        End If
            
        If Not blnSkipWidthAdjust Then
            If grdIonList.Visible Then
                .Width = Me.Width - .Left - 220 - grdIonList.Width - 90
            Else
                .Width = Me.Width - .Left - 220
            End If
        End If
        .Height = lngPreferredHeight
    End With
    
    mnuIonMatchListRightClick.Visible = False
    
End Sub

Private Function QueryExistingIonsInList() As Boolean
    Dim strMessage As String, eResponse As VbMsgBoxResult
    
    If grdIonList.Rows > 1 Then
        strMessage = LookupMessage(910)
        eResponse = MsgBox(strMessage, vbQuestion + vbYesNoCancel + vbDefaultButton3, LookupMessage(920))
        If eResponse = vbYes Then
            QueryExistingIonsInList = True
        Else
            QueryExistingIonsInList = False
        End If
    Else
        QueryExistingIonsInList = True
    End If
    
End Function

Public Sub ResetCWSpectrumOptions()
    objSpectrum.ResetOptionsToDefaults True, True, 2, pmStickToZero
End Sub

Public Sub ResetPredictedSpectrumCRC()
    CRCForIonListSaved = 0
    CRCForPredictedSpectrumSaved = 0
End Sub

Public Sub SaveCWSpectrumOptions(strSequenceFilePath As String)
On Error GoTo SaveCWSpectrumOptionsErrorHandler
    
    objSpectrum.SaveDataToDisk strSequenceFilePath, True, ",", False, True
    Exit Sub

SaveCWSpectrumOptionsErrorHandler:
    Debug.Assert False
    GeneralErrorHandler "frmFragmentationModelling|SaveDataToDisk", Err.Number, Err.Description
End Sub

Private Function SearchForIonInIonListGrid(ByVal dblMassToFind As Double, ByVal dblMatchTolerance As Double) As Long
    ' Returns -1 if no match
    ' Otherwise, returns the row index of the match in grdIonList
    
    Dim lngIonIndex As Long, boolIonMassMatched As Boolean
    Dim dblCompareMass As Double
    
    boolIonMassMatched = False
    For lngIonIndex = 1 To grdIonList.Rows - 1
        dblCompareMass = grdIonList.TextMatrix(lngIonIndex, ilgMass)
        If WithinTolerance(dblMassToFind, dblCompareMass, dblMatchTolerance) Then
            boolIonMassMatched = True
            Exit For
        End If
    Next lngIonIndex
    
    If boolIonMassMatched Then
        SearchForIonInIonListGrid = lngIonIndex
    Else
        SearchForIonInIonListGrid = -1
    End If
End Function

Private Sub SetDtaTxtFileBrowserMenuVisibility(blnShowMenu As Boolean)
    mnuViewDtaTxtSep.Visible = blnShowMenu
    mnuViewDtaTxtBrowser.Visible = blnShowMenu
End Sub

Public Sub SetIonMatchList(dblXVals() As Double, dblYVals() As Double, lngDataCount As Long, strSourceFilePath As String, lngScanNumberStart As Long, lngScanNumberEnd As Long, ByVal dblParentIon As Double, intParentIonCharge As Integer, Optional blnAllowAutoBinningDetermination As Boolean = True)
    ' Set the ion match list values via code
    ' dblXVals() and dblYVals() should be 0-based
    
    Dim lngIndex As Long
    Dim dblMinMass As Double, dblMaxMass As Double
    Dim strCaption As String, strPrecursorIonMH As String
    Dim fso As New FileSystemObject
    
    ' Initialize IonMatchList
    InitializeIonMatchList lngDataCount
    
    dblMinMass = HighestValueForDoubleDataType
    dblMaxMass = LowestValueForDoubleDataType

    For lngIndex = 0 To lngDataCount - 1
        AppendXYPair IonMatchList(), IonMatchListCount, dblXVals(lngIndex), dblYVals(lngIndex), True
        
        If dblXVals(lngIndex) < dblMinMass Then dblMinMass = dblXVals(lngIndex)
        If dblXVals(lngIndex) > dblMaxMass Then dblMaxMass = dblXVals(lngIndex)
    Next lngIndex

    If lngDataCount > 0 Then
        If Abs((dblMaxMass - dblMinMass) / lngDataCount) < ION_SEPARATION_TOLERANCE_TO_AUTO_ENABLE_BINNING Then
            ' Average spacing between ions is less than 0.2 Da, turn on binning
            frmIonMatchOptions.chkGroupSimilarIons.value = vbChecked
        End If
        
        ' Reset the alignment offset
        txtAlignment = "0"
            
        ' Note that the Parent Ion mass in a .Dta file is the MH mass, and is thus already 1+
        strPrecursorIonMH = FormatNumberAsString(Round(dblParentIon, 2), 10, 8)
        
        strCaption = CompactPathString(fso.GetFileName(strSourceFilePath), 20)
        If Right(strCaption, 2) = ".." Then strCaption = Left(strCaption, Len(strCaption) - 2)
        
        strCaption = strCaption & ": Scan " & Trim(lngScanNumberStart)
        If lngScanNumberEnd <> lngScanNumberStart Then
            strCaption = strCaption & "-" & Trim(lngScanNumberEnd)
        End If
    
        mIonMatchListDetails.Caption = strCaption
        
        txtPrecursorIonMass = strPrecursorIonMH
        
        UpdateIonMatchListWrapper
        
    End If
    
    Set fso = Nothing
    
End Sub

Public Sub SetIonMatchListCaption(strNewCaption As String)
    mIonMatchListDetails.Caption = strNewCaption
End Sub

Private Sub SetIonMatchListVisibility(blnShowList As Boolean)
    mnuViewMatchIonList.Checked = blnShowList
    
    grdIonList.Visible = blnShowList
    mnuClearMatchIonList.Enabled = blnShowList
    fraIonMatching.Visible = blnShowList
    fraIonStats.Visible = blnShowList
    cmdMatchIons.Visible = blnShowList
    
    If Me.WindowState = vbNormal And blnShowList Then
        If Me.Width < 9000 Then Me.Width = 9000
    End If
    
    PositionControls
    UpdateMassesGridAndSpectrumWrapper
    
End Sub

Public Sub SetFragMatchSettingsChanged()
    mFragMatchSettingsChanged = True
End Sub

Public Sub SetNeedToZoomOutFull(blnEnable As Boolean)
    mNeedToZoomOutFull = blnEnable
End Sub

Public Sub SetUpdateSpectrumMode(blnUpdateOnChange As Boolean)
    mnuUpdateSpectrum.Checked = blnUpdateOnChange
End Sub
    
Private Sub ShellSortIonList(ByRef ThisIonMatchList() As Double, lngPointerArray() As Long, ByVal lngLowIndex As Long, ByVal lngHighIndex As Long, intColIndexToSortOn As Integer, blnUpdateProgressForm As Boolean)
    ' Sorts the PointerArray to ThisIonMatchList() using column intColIndexToSortOn
    ' Note that ThisIonMatchList() is a 2D array, 1-based in the first dimension but uses columns 0, 1, and 2
    
    Dim lngCount As Long
    Dim lngIncrement As Long
    Dim lngIndex As Long
    Dim lngIndexCompare As Long
    Dim lngPointerSaved As Long
    Dim dblCompareVal As Double

    ' sort array[lngLowIndex..lngHighIndex]
    
    ' compute largest increment
    lngCount = lngHighIndex - lngLowIndex + 1
    lngIncrement = 1
    If (lngCount < 14) Then
        lngIncrement = 1
    Else
        Do While lngIncrement < lngCount
            lngIncrement = 3 * lngIncrement + 1
        Loop
        lngIncrement = lngIncrement \ 3
        lngIncrement = lngIncrement \ 3
    End If

    Do While lngIncrement > 0
        ' sort by insertion in increments of lngIncrement
        For lngIndex = lngLowIndex + lngIncrement To lngHighIndex
            lngPointerSaved = lngPointerArray(lngIndex)
            dblCompareVal = ThisIonMatchList(lngPointerSaved, intColIndexToSortOn)
            For lngIndexCompare = lngIndex - lngIncrement To lngLowIndex Step -lngIncrement
                ' Use <= to sort ascending; Use > to sort descending
                If ThisIonMatchList(lngPointerArray(lngIndexCompare), intColIndexToSortOn) > dblCompareVal Then Exit For
                lngPointerArray(lngIndexCompare + lngIncrement) = lngPointerArray(lngIndexCompare)
            Next lngIndexCompare
            lngPointerArray(lngIndexCompare + lngIncrement) = lngPointerSaved
        Next lngIndex
        lngIncrement = lngIncrement \ 3
    Loop

End Sub
    
Public Sub ShowAutoLabelPeaksOptions()
    objSpectrum.ShowAutoLabelPeaksDialog
End Sub

Private Sub UpdateFragmentationSpectrumOptions()
    
    Dim intIonIndex As Integer
    Dim blnModifyIon As Boolean
    Dim sngDivisorForInversion As Single
    Dim intLastGoodLineNumber As Integer

On Error GoTo UpdateFragmentationSpectrumOptionsHandler
    intLastGoodLineNumber = 2568
    
    If objMwtWin.Peptide Is Nothing Then Exit Sub
        
    intLastGoodLineNumber = 2572
    ' Initialize to the current values
    udtNewFragmentationSpectrumOptions = objMwtWin.Peptide.GetFragmentationSpectrumOptions
    
    intLastGoodLineNumber = 2576
    With udtNewFragmentationSpectrumOptions
        intLastGoodLineNumber = 2578
        .DoubleChargeIonsShow = cChkBox(chkDoubleCharge)
        .TripleChargeIonsShow = cChkBox(chkTripleCharge)
        
        If cboDoubleCharge.ListIndex >= 0 And cboDoubleCharge.ListCount > 0 Then
            .DoubleChargeIonsThreshold = cboDoubleCharge.List(cboDoubleCharge.ListIndex)
        End If
        
        If cboTripleCharge.ListIndex >= 0 And cboTripleCharge.ListCount > 0 Then
            .TripleChargeIonsThreshold = cboTripleCharge.List(cboTripleCharge.ListIndex)
        End If
        
        intLastGoodLineNumber = 2584
        For intIonIndex = 0 To TOTAL_POSSIBLE_ION_TYPES - 1
            .IonTypeOptions(intIonIndex).ShowIon = cChkBox(chkIonType(intIonIndex))
        Next intIonIndex
        
        intLastGoodLineNumber = 2589
        For intIonIndex = 0 To TOTAL_POSSIBLE_ION_TYPES - 1
            blnModifyIon = lstIonsToModify.Selected(intIonIndex)
            
            .IonTypeOptions(intIonIndex).NeutralLossWater = blnModifyIon And cChkBox(chkWaterLoss)
            .IonTypeOptions(intIonIndex).NeutralLossAmmonia = blnModifyIon And cChkBox(chkAmmoniaLoss)
            .IonTypeOptions(intIonIndex).NeutralLossPhosphate = blnModifyIon And cChkBox(chkPhosphateLoss)
        
        Next intIonIndex
    
        ' Note: A ions can have ammonia and phosphate loss, but not water loss, so always set this to false
        .IonTypeOptions(itAIon).NeutralLossWater = False
    
        intLastGoodLineNumber = 2602
        For intIonIndex = 0 To TOTAL_POSSIBLE_ION_TYPES - 1
            .IntensityOptions.IonType(intIonIndex) = Val(frmIonMatchOptions.txtIonIntensity(intIonIndex))
        Next intIonIndex
        .IntensityOptions.BYIonShoulder = frmIonMatchOptions.txtBYIonShoulders
        .IntensityOptions.NeutralLoss = frmIonMatchOptions.txtNeutralLosses
    
        intLastGoodLineNumber = 2609
        If cChkBox(frmIonMatchOptions.chkPlotSpectrumInverted) Then
            If cChkBox(frmIonMatchOptions.chkFragSpecEmphasizeProlineYIons) Then
                sngDivisorForInversion = 2.5
            Else
                sngDivisorForInversion = 2.77777
            End If
            
            For intIonIndex = 0 To TOTAL_POSSIBLE_ION_TYPES - 1
                .IntensityOptions.IonType(intIonIndex) = -Abs(.IntensityOptions.IonType(intIonIndex)) / sngDivisorForInversion
            Next intIonIndex
            .IntensityOptions.BYIonShoulder = -Abs(.IntensityOptions.BYIonShoulder) / sngDivisorForInversion
            .IntensityOptions.NeutralLoss = -Abs(.IntensityOptions.NeutralLoss) / sngDivisorForInversion
        End If
    
    End With
    
    intLastGoodLineNumber = 2626
    objMwtWin.Peptide.SetFragmentationSpectrumOptions udtNewFragmentationSpectrumOptions
    Exit Sub
    
UpdateFragmentationSpectrumOptionsHandler:
Resume
    GeneralErrorHandler "frmFragmentationModelling|UpdateFragmentationSpectrumOptions", Err.Number, "Last good line number = " & Trim(intLastGoodLineNumber) & vbCrLf & Err.Description
    
End Sub

Private Sub UpdateIonLoadStats(Optional boolClearStats As Boolean = False)
    Dim lngWithinTolerance As Long, sngPercentage As Single
    Dim strLoaded As String, strRemaining As String
    Dim strWithinTolerance As String
    Dim strMessage As String
    
    strLoaded = LookupLanguageCaption(12370, "Loaded") & ": "
    strRemaining = LookupLanguageCaption(12375, "Remaining after binning") & ": "
    strWithinTolerance = LookupLanguageCaption(12380, "Within tolerance") & ": "
    
    lblIonLoadedCount.Caption = strLoaded & CStr(IonMatchListCount)
    
    If boolClearStats Then
        lblBinAndToleranceCounts.Caption = ""
        lblPrecursorStatus.Caption = ""
        lblMatchCount.Caption = ""
        lblScore.Caption = ""
    Else
        lngWithinTolerance = grdIonList.Rows - 1
        If BinnedDataCount > 0 Then
            strMessage = strRemaining & CStr(BinnedDataCount) & vbCrLf
            sngPercentage = lngWithinTolerance / BinnedDataCount * 100
        Else
            If IonMatchListCount > 0 Then
                sngPercentage = lngWithinTolerance / IonMatchListCount * 100
            Else
                sngPercentage = 0
            End If
        End If
        strMessage = strMessage & strWithinTolerance & CStr(lngWithinTolerance) & " (" & Format(sngPercentage, "#0.0") & "%)"
        lblBinAndToleranceCounts = strMessage
    End If

End Sub

Public Sub UpdateIonMatchListWrapper()
    Dim blnUpdateMassSpectrumSaved As Boolean
    
    blnUpdateMassSpectrumSaved = mnuUpdateSpectrum.Checked
    If blnUpdateMassSpectrumSaved Then mnuUpdateSpectrum.Checked = False
    
    ' The following calls UpdateMassesGridAndSpectrumWrapper
    SetIonMatchListVisibility True
    BinIonMatchList
    
    If blnUpdateMassSpectrumSaved Then mnuUpdateSpectrum.Checked = True
    
    ' The following sub calls UpdateMassesGridAndSpectrumWrapper, which calls the MatchIons sub
    NormalizeIonMatchListWrapper

    UpdateIonLoadStats

End Sub

Private Sub UpdateMassesGridAndSpectrumWrapper()
    
    ' Set the Settings Changed Bit
    SetFragMatchSettingsChanged
    
    ' Display predicted ions & intensities in grid
    DisplayPredictedIonMasses
    MatchIons
    
    ' Call UpdatePredictedFragSpectrum if necessary
    If mnuUpdateSpectrum.Checked Then
        DisplaySpectra
    End If
End Sub

Private Sub UpdateMatchCountAndScoreWork(lngMatchCount As Long, dblMatchScore As Double)
    lblMatchCount = LookupLanguageCaption(12400, "Matches") & ": " & CStr(lngMatchCount)
    lblScore = LookupLanguageCaption(12405, "Score") & ": " & Format(dblMatchScore, "0.0")
End Sub

Private Sub UpdatePredictedFragMasses()
    ' Determines the masses of the expected ions for the given sequence
    
    Dim dblSequenceMass As Double
    
    If mDelayUpdate Then Exit Sub
    
    If Len(txtSequence) > 100 Then Me.MousePointer = vbHourglass
    
    objMwtWin.Peptide.SetSequence txtSequence, GetNTerminusState(), GetCTerminusState(), Get3LetterCodeState()
    
    If objMwtWin.Peptide.GetResidueCount > 0 Then
        dblSequenceMass = objMwtWin.Peptide.GetPeptideMass()
        If dblSequenceMass = 0 Then
            txtMH = 0
        Else
            If cboNTerminus.ListIndex = ntgHydrogenPlusProton Then
                ' Don't need to add a proton
                txtMH = Round(dblSequenceMass, 6)
            Else
                txtMH = Round(dblSequenceMass + objMwtWin.GetChargeCarrierMass, 6)
            End If
        End If
    Else
        dblSequenceMass = 0
        txtMH = "0"
    End If
    
    txtMWT = LookupLanguageCaption(4040, "MW") & " = " & Round(dblSequenceMass, 6)
    
    UpdateMassesGridAndSpectrumWrapper
    
    If Len(txtSequence) > 100 Then Me.MousePointer = vbDefault
End Sub

Private Sub UpdateStandardMasses()
    
    dblHistidineFW = objMwtWin.ComputeMass("His")
    dblPhenylalanineFW = objMwtWin.ComputeMass("Phe")
    dblTyrosineFW = objMwtWin.ComputeMass("Tyr")
    dblImmoniumMassDifference = objMwtWin.ComputeMass("CO") - objMwtWin.ComputeMass("H")
   
    UpdatePredictedFragMasses
End Sub

Private Function WithinTolerance(ThisNumber As Double, CompareNumber As Double, ThisTolerance As Double) As Boolean
    If Abs(ThisNumber - CompareNumber) <= ThisTolerance Then
        WithinTolerance = True
    Else
        WithinTolerance = False
    End If
End Function

Private Sub cboCTerminus_Click()
    UpdatePredictedFragMasses
End Sub

Private Sub cboDoubleCharge_Click()
    UpdatePredictedFragMasses
End Sub

Private Sub cboMHAlternate_Click()
    ConvertSequenceMH True
End Sub

Private Sub cboNotation_Click()
    UpdatePredictedFragMasses
End Sub

Private Sub cboNTerminus_Click()
    UpdatePredictedFragMasses
End Sub

Private Sub cboTripleCharge_Click()
    UpdatePredictedFragMasses
End Sub

Private Sub chkAmmoniaLoss_Click()
    UpdateMassesGridAndSpectrumWrapper
End Sub

Private Sub chkDoubleCharge_Click()
    EnableDisableControls
    UpdateMassesGridAndSpectrumWrapper
End Sub

Private Sub chkPhosphateLoss_Click()
    UpdateMassesGridAndSpectrumWrapper
End Sub

Private Sub chkRemovePrecursorIon_Click()
    EnableDisableControls
    UpdateIonMatchListWrapper
End Sub

Private Sub chkTripleCharge_Click()
    EnableDisableControls
    UpdateMassesGridAndSpectrumWrapper
End Sub

Private Sub chkWaterLoss_Click()
    UpdateMassesGridAndSpectrumWrapper
End Sub

Private Sub chkIonType_Click(Index As Integer)
    UpdateMassesGridAndSpectrumWrapper
End Sub

Private Sub cmdMatchIons_Click()
    MatchIons
End Sub

Private Sub Form_Activate()
    Dim lclFragMatchSettingsChanged As Boolean
    
    PossiblyHideMainWindow
    
    lclFragMatchSettingsChanged = mFragMatchSettingsChanged
    mFormActivatedByUser = True
    
    Select Case objMwtWin.GetElementMode
    Case emAverageMass
        optElementMode(0).value = True
    Case emIsotopicMass
        optElementMode(1).value = True
    Case Else
        objMwtWin.SetElementMode emIsotopicMass
        optElementMode(1).value = True
    End Select
    
    UpdateStandardMasses
    UpdateIonLoadStats
    
    If lclFragMatchSettingsChanged And IonMatchListCount > 0 Then
        ' Does this code ever get reached?
        ' Yes, if no ion types are selected, then this code could be encountered
        Debug.Assert False
        UpdateIonMatchListWrapper
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        mnuClose_Click
    End If
    
End Sub

Private Sub Form_Load()
    Dim strCurrentTask As String
    
On Error GoTo FormLoadErrorHandler

    strCurrentTask = "Loading frmFragmentationModelling"
    
    If objSpectrum Is Nothing Then
        strCurrentTask = "Instantiating CWSpectrumDLL.Spectrum on frmFragmentationModelling"
        Set objSpectrum = New CWSpectrumDLL.Spectrum
    End If
    
    strCurrentTask = "Initializing objSpectrum on frmFragmentationModelling"
    With objSpectrum
        .SetSeriesCount 2
        .SetSeriesPlotMode 1, pmStickToZero, True
    End With
    mNeedToZoomOutFull = True
    
    strCurrentTask = "Initializing frmFragmentationModelling"
    
    SetDtaTxtFileBrowserMenuVisibility False
    SetIonMatchListVisibility False
    
    PopulateComboBoxes
    
    ' Note that PositionControls gets called when the form is resized with the following command
    ' Must call SetIonMatchListVisibility before calling this sub
    SizeAndCenterWindow Me, cWindowTopLeft, 12000, 9250
    
    EnableDisableControls
    
    UpdateStandardMasses
    
    InitializeIonMatchList 0
    
    InitializeIonListGrid
    UpdateMassesGridAndSpectrumWrapper
    
    Exit Sub

FormLoadErrorHandler:
    GeneralErrorHandler "frmFragmentationModelling|Form_Load", Err.Number, "Error " & strCurrentTask & ": " & Err.Description
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    QueryUnloadFormHandler Me, Cancel, UnloadMode
End Sub

Private Sub Form_Resize()
    PositionControls
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objSpectrum = Nothing
End Sub

Private Sub grdFragMasses_KeyDown(KeyCode As Integer, Shift As Integer)
    FlexGridKeyPressHandler Me, grdFragMasses, KeyCode, Shift
End Sub

Private Sub grdIonList_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyV And (Shift And vbCtrlMask) Then
        PasteIonMatchList
    Else
        FlexGridKeyPressHandler Me, grdIonList, KeyCode, Shift
    End If
        
End Sub

Private Sub grdIonList_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        Me.PopupMenu mnuIonMatchListRightClick
    End If
End Sub

Private Sub lblIonMatchingWindow_Click()
    Static ClickTrack As Integer, eResponse As VbMsgBoxResult
    
    ClickTrack = ClickTrack + 1
    
    If ClickTrack = 3 Then
        ClickTrack = 0
        eResponse = MsgBox("Initialize with Dummy Data?", vbQuestion + vbYesNoCancel, "Initialize")
        If eResponse = vbYes Then InitializeDummyData 2
    End If
End Sub

Private Sub lstIonsToModify_Click()
    UpdateMassesGridAndSpectrumWrapper
End Sub

Private Sub mniIonMatchListRightClickPaste_Click()
    PasteIonMatchList
End Sub

Private Sub mnuAutoAlign_Click()
    AutoAlignMatchIonList
End Sub

Private Sub mnuClearMatchIonList_Click()
    IonMatchListClear
End Sub

Private Sub mnuClose_Click()
    HideFormShowMain Me
End Sub

Private Sub mnuCopySequenceMW_Click()
    CopySequenceMW
End Sub

Private Sub mnuCopyPredictedIons_Click()
    CopyFragGridInfo grdFragMasses, gcmText
End Sub

Private Sub mnuCopyPredictedIonsAsRTF_Click()
    CopyFragGridInfo grdFragMasses, gcmRTF
End Sub

Private Sub mnuCopyPredictedIonsAsHtml_Click()
    CopyFragGridInfo grdFragMasses, gcmHTML
End Sub

Private Sub mnuEditModificationSymbols_Click()
    frmAminoAcidModificationSymbols.InitializeForm
    frmAminoAcidModificationSymbols.Show vbModal
    UpdatePredictedFragMasses
End Sub
    
Private Sub mnuFragmentationModellingHelp_Click()
    ShowHelpPage hwnd, 3080

End Sub

Private Sub mnuIonMatchListOptions_Click()
    frmIonMatchOptions.Show vbModeless
End Sub

Private Sub mnuIonMatchListRightClickCopy_Click()
    ' Equivalent to pressing Ctrl+C
    FlexGridCopyInfo Me, grdIonList, gcmText
End Sub

Private Sub mnuIonMatchListRightClickDeleteAll_Click()
    IonMatchListClear
End Sub

Private Sub mnuIonMatchListRightClickSelectAll_Click()
    ' Equivalent to pressing Ctrl+A
    FlexGridSelectEntireGrid grdIonList
End Sub

Private Sub mnuLoadIonList_Click()
    Dim blnSuccess As Boolean
    
'''    ' Warn user about replacing data
'''    If Not QueryExistingIonsInList() Then Exit Sub
    
    blnSuccess = LoadIonListToMatch()
    
    If blnSuccess Then
        UpdateIonMatchListWrapper
    Else
        If Err.Number <> 32755 Then
            MsgBox LookupMessage(1020), vbInformation + vbOKOnly, LookupMessage(985)
        End If
    End If
    
End Sub

Private Sub mnuLoadSequenceInfo_Click()
    LoadSequenceInfo
End Sub

Private Sub mnuPasteIonList_Click()
    PasteIonMatchList
End Sub

Private Sub mnuSaveSequenceInfo_Click()
    SaveSequenceInfo IonMatchList(), IonMatchListCount, mIonMatchListDetails.Caption
End Sub

Private Sub mnuShowMassSpectrum_Click()
    SetUpdateSpectrumMode True
    
    UpdateMassesGridAndSpectrumWrapper
    
    objSpectrum.ShowSpectrum
End Sub

Private Sub mnuViewDtaTxtBrowser_Click()
    If frmDtaTxtFileBrowser.GetDataInitializedState() Then
        frmDtaTxtFileBrowser.Show
    End If
End Sub

Private Sub mnuViewMatchIonList_Click()
    SetIonMatchListVisibility (Not mnuViewMatchIonList.Checked)
End Sub

Private Sub mnuUpdateSpectrum_Click()
    SetUpdateSpectrumMode Not mnuUpdateSpectrum.Checked
    UpdateMassesGridAndSpectrumWrapper
End Sub

Private Sub objSpectrum_SpectrumFormRequestClose()
    ' The SpectrumForm was closed (actually, most likely just hidden)
    ' If we wanted to do anything special, we could do it here
End Sub

Private Sub optElementMode_Click(Index As Integer)
    ' Elementweightmode = 1 means average weights, 2 means isotopic, and 3 means integer
    Dim eNewWeightMode As emElementModeConstants
    
    eNewWeightMode = Index + 1
        
    If eNewWeightMode <> objMwtWin.GetElementMode Then
        SwitchWeightMode eNewWeightMode, True, False
        
        UpdateStandardMasses
        UpdateIonLoadStats
    
        If mFragMatchSettingsChanged And IonMatchListCount > 0 Then
            ' Does this code ever get reached?
            Debug.Assert False
            UpdateIonMatchListWrapper
        End If
    End If
    
End Sub

Private Sub txtAlignment_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And Not mAutoAligning Then
        AlignmentOffsetValidate
    End If
End Sub

Private Sub txtAlignment_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtAlignment, KeyAscii, True, True, True, False, True, False, False, False, False, True
End Sub

Private Sub txtAlignment_Validate(Cancel As Boolean)
    If Val(GetMostRecentTextBoxValue) <> Val(txtAlignment) Then
        If Not mAutoAligning Then
            AlignmentOffsetValidate
        End If
    End If
End Sub

Private Sub txtIonMatchingWindow_GotFocus()
    HighlightOnFocus txtIonMatchingWindow
End Sub

Private Sub txtIonMatchingWindow_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then UpdateIonMatchListWrapper
End Sub

Private Sub txtIonMatchingWindow_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtIonMatchingWindow, KeyAscii, True, True, False
End Sub

Private Sub txtIonMatchingWindow_Validate(Cancel As Boolean)
    ValidateTextboxValueDbl txtIonMatchingWindow, 0, 100, 0.5
    
    If Val(GetMostRecentTextBoxValue) <> Val(txtIonMatchingWindow) Then UpdateIonMatchListWrapper

End Sub

Private Sub txtMH_Change()
    ConvertSequenceMH True
End Sub

Private Sub txtMH_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtMH, KeyAscii, True, True, True, False, True, False, False, False, False, True, True
End Sub

Private Sub txtMH_Validate(Cancel As Boolean)
    If IsNumeric(txtMH) Then
        If CDblSafe(txtMH) < 0 Then txtMH = "0"
    End If
End Sub

Private Sub txtMHAlt_Change()
    ConvertSequenceMH False
End Sub

Private Sub txtMHAlt_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtMHAlt, KeyAscii, True, True, True, False, True, False, False, False, False, True, True
End Sub

Private Sub txtMHAlt_Validate(Cancel As Boolean)
    If IsNumeric(txtMH) Then
        If CDblSafe(txtMH) < 0 Then txtMH = "0"
    End If
End Sub

Private Sub txtPrecursorIonMass_GotFocus()
    HighlightOnFocus txtPrecursorIonMass
End Sub

Private Sub txtPrecursorIonMass_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then UpdateIonMatchListWrapper
End Sub

Private Sub txtPrecursorIonMass_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtPrecursorIonMass, KeyAscii, True, True, False
End Sub

Private Sub txtPrecursorIonMass_Validate(Cancel As Boolean)
    ValidateTextboxValueDbl txtPrecursorIonMass, 0, 32000, 500
    
    If Val(GetMostRecentTextBoxValue) <> Val(txtPrecursorIonMass) Then UpdateIonMatchListWrapper
    
End Sub

Private Sub txtPrecursorMassWindow_GotFocus()
    HighlightOnFocus txtPrecursorMassWindow
End Sub

Private Sub txtPrecursorMassWindow_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then UpdateIonMatchListWrapper
End Sub

Private Sub txtPrecursorMassWindow_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtPrecursorMassWindow, KeyAscii, True, True
End Sub

Private Sub txtPrecursorMassWindow_Validate(Cancel As Boolean)
    ValidateTextboxValueDbl txtPrecursorMassWindow, 0, 100, 2
    
    If Val(GetMostRecentTextBoxValue) <> Val(txtPrecursorMassWindow) Then UpdateIonMatchListWrapper

End Sub

Private Sub txtSequence_Change()
    If mDelayUpdate Then Exit Sub
    
    CheckSequenceTerminii
    
    UpdatePredictedFragMasses
End Sub

Private Sub txtSequence_GotFocus()
    HighlightOnFocus txtSequence
End Sub

Private Sub txtSequence_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strText As String
    
    If Shift And vbCtrlMask Then
        If KeyCode = vbKeyC Then
            ' Ctrl+C was used
            ' For some reason automatic handling of Ctrl+C isn't working with txtSequence
            ' Thus, manually copy to clipboard
            
            strText = Mid(txtSequence.Text, txtSequence.SelStart + 1, txtSequence.SelLength)
            Clipboard.SetText strText, vbCFText
        End If
    End If
End Sub

Private Sub txtSequence_KeyPress(KeyAscii As Integer)
    PeptideSequenceKeyPressHandler txtSequence, KeyAscii
End Sub
