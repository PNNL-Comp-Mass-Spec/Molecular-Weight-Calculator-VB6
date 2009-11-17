VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmMMConvert 
   Caption         =   "Mole/Mass Converter and Dilution Calculator"
   ClientHeight    =   6660
   ClientLeft      =   540
   ClientTop       =   1395
   ClientWidth     =   10515
   ForeColor       =   &H80000008&
   HelpContextID   =   3040
   Icon            =   "MMCONV.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6660
   ScaleWidth      =   10515
   Tag             =   "11000"
   Begin VB.Frame fraMZConversion 
      Height          =   1770
      Left            =   120
      TabIndex        =   13
      Tag             =   "12410"
      Top             =   4750
      Width           =   6315
      Begin VB.TextBox txtMZHighAlt 
         Height          =   315
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   31
         Tag             =   "11100"
         Text            =   "100"
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox txtMZLowAlt 
         Height          =   315
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   30
         Tag             =   "11100"
         Text            =   "100"
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox txtMZHigh 
         Height          =   315
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   29
         Tag             =   "11100"
         Text            =   "100"
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtMZLow 
         Height          =   315
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   28
         Tag             =   "11100"
         Text            =   "100"
         Top             =   960
         Width           =   1095
      End
      Begin VB.ComboBox cboMassError 
         Height          =   315
         Left            =   5040
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Tag             =   "11020"
         Top             =   240
         Width           =   1125
      End
      Begin VB.TextBox txtMassError 
         Height          =   315
         Left            =   4200
         TabIndex        =   24
         Tag             =   "11100"
         Text            =   "100"
         Top             =   240
         Width           =   735
      End
      Begin VB.ComboBox cboMH 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   960
         Width           =   1600
      End
      Begin VB.TextBox txtMH 
         Height          =   315
         Left            =   1920
         TabIndex        =   18
         Text            =   "1000"
         Top             =   960
         Width           =   1200
      End
      Begin VB.ComboBox cboMHAlternate 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   1320
         Width           =   1600
      End
      Begin VB.TextBox txtMHAlt 
         Height          =   315
         Left            =   1920
         TabIndex        =   21
         Text            =   "500.503695"
         Top             =   1320
         Width           =   1200
      End
      Begin VB.OptionButton optElementMode 
         Caption         =   "Isotopic"
         Height          =   255
         HelpContextID   =   4053
         Index           =   1
         Left            =   1080
         TabIndex        =   16
         Tag             =   "12430"
         Top             =   480
         Width           =   1395
      End
      Begin VB.OptionButton optElementMode 
         Caption         =   "Average"
         Height          =   255
         HelpContextID   =   4053
         Index           =   0
         Left            =   1080
         TabIndex        =   15
         Tag             =   "12425"
         Top             =   240
         Value           =   -1  'True
         Width           =   1395
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "m/z end"
         Height          =   255
         Left            =   5040
         TabIndex        =   27
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "m/z start"
         Height          =   255
         Left            =   3840
         TabIndex        =   26
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblMassError 
         Caption         =   "Mass Error"
         Height          =   255
         Left            =   3240
         TabIndex        =   23
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblMHDa 
         Caption         =   "m/z"
         Height          =   255
         Left            =   3200
         TabIndex        =   19
         Tag             =   "12350"
         Top             =   985
         Width           =   495
      End
      Begin VB.Label lblMHAltDa 
         Caption         =   "m/z"
         Height          =   255
         Left            =   3200
         TabIndex        =   22
         Tag             =   "12350"
         Top             =   1345
         Width           =   495
      End
      Begin VB.Label lblElementMode 
         Caption         =   "Element Mode"
         Height          =   465
         Left            =   120
         TabIndex        =   14
         Tag             =   "12420"
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.TextBox txtDensity 
      Height          =   285
      Left            =   3240
      TabIndex        =   61
      Tag             =   "11110"
      Text            =   "1"
      ToolTipText     =   "Density of compound"
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CheckBox chkLinkDilutionVolumeUnits 
      Caption         =   "Link Dilution Volume Units"
      Height          =   405
      HelpContextID   =   4092
      Left            =   6240
      TabIndex        =   33
      Tag             =   "11230"
      Top             =   1080
      Width           =   4095
   End
   Begin VB.CheckBox chkLinkMolarities 
      Caption         =   "&Link Initial Dilution Concentration and Convert Amounts Concentration"
      Height          =   405
      HelpContextID   =   4092
      Left            =   6240
      TabIndex        =   32
      Tag             =   "11220"
      Top             =   600
      Width           =   3975
   End
   Begin VB.Frame fraConvertAmounts 
      Height          =   2445
      Left            =   120
      TabIndex        =   1
      Top             =   2040
      Width           =   4935
      Begin VB.TextBox txtConcentration 
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Tag             =   "11130"
         Text            =   "1"
         ToolTipText     =   "Molarity of compound in the solvent"
         Top             =   1920
         Width           =   1095
      End
      Begin VB.ComboBox cboVolume 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Tag             =   "11030"
         ToolTipText     =   "Units of volume"
         Top             =   1320
         Width           =   1600
      End
      Begin VB.TextBox txtVolume 
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Tag             =   "11120"
         Text            =   "100"
         ToolTipText     =   "Volume of solvent the compound is dissolved in"
         Top             =   1320
         Width           =   1095
      End
      Begin VB.ComboBox cboTo 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Tag             =   "11020"
         Top             =   720
         Width           =   1600
      End
      Begin VB.ComboBox cboFrom 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Tag             =   "11020"
         Top             =   240
         Width           =   1600
      End
      Begin VB.TextBox txtFromNum 
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Tag             =   "11100"
         Text            =   "1"
         ToolTipText     =   "Amount of compound to convert from or to be dissolved in solvent"
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdFindVolume 
         Caption         =   "Find &Volume"
         Height          =   480
         Left            =   3075
         TabIndex        =   9
         Tag             =   "11160"
         ToolTipText     =   "Calculate the volume using the amount and molarity"
         Top             =   1200
         Width           =   1515
      End
      Begin VB.CommandButton cmdFindConcentration 
         Caption         =   "Find &Concentration"
         Height          =   480
         Left            =   3075
         TabIndex        =   12
         Tag             =   "11170"
         ToolTipText     =   "Calculate the concentration using the amount and volume"
         Top             =   1800
         Width           =   1515
      End
      Begin VB.ComboBox cboConcentration 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Tag             =   "7090"
         Top             =   1920
         Width           =   1600
      End
      Begin VB.CommandButton cmdFindAmount 
         Caption         =   "Find &Amount"
         Height          =   480
         Left            =   3075
         TabIndex        =   6
         Tag             =   "11150"
         ToolTipText     =   "Calculate the amount using the volume and molarity"
         Top             =   240
         Width           =   1515
      End
      Begin VB.Label lblToNum 
         Caption         =   "Label1"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.Frame fraDilutionCalcs 
      Caption         =   "Dilution Calculations"
      Height          =   3135
      Left            =   5160
      TabIndex        =   34
      Tag             =   "11200"
      Top             =   1560
      Width           =   5175
      Begin VB.ComboBox cboDilutionMode 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Tag             =   "11210"
         ToolTipText     =   "Quantity to find for dilution calculations"
         Top             =   240
         Width           =   3675
      End
      Begin VB.ComboBox cboDilutingSolventVolume 
         Height          =   315
         Left            =   3360
         Style           =   2  'Dropdown List
         TabIndex        =   47
         Tag             =   "11030"
         Top             =   2160
         Width           =   1600
      End
      Begin VB.TextBox txtDilutingSolventVolume 
         Height          =   315
         Left            =   2160
         TabIndex        =   46
         Tag             =   "11285"
         Text            =   "12"
         ToolTipText     =   "Volume of solvent the compound is dissolved in"
         Top             =   2160
         Width           =   1015
      End
      Begin VB.ComboBox cboStockSolutionVolume 
         Height          =   315
         Left            =   3360
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Tag             =   "11030"
         Top             =   1200
         Width           =   1600
      End
      Begin VB.TextBox txtStockSolutionVolume 
         Height          =   315
         Left            =   2160
         TabIndex        =   40
         Tag             =   "11265"
         Text            =   "3"
         ToolTipText     =   "Volume of solvent the compound is dissolved in"
         Top             =   1200
         Width           =   1015
      End
      Begin VB.TextBox txtDilutionConcentrationInitial 
         Height          =   315
         Left            =   2160
         TabIndex        =   37
         Tag             =   "11255"
         Text            =   "10"
         ToolTipText     =   "Molarity of compound in the solvent"
         Top             =   720
         Width           =   1015
      End
      Begin VB.ComboBox cboDilutionConcentrationInitial 
         Height          =   315
         Left            =   3360
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Tag             =   "7090"
         Top             =   720
         Width           =   1600
      End
      Begin VB.ComboBox cboDilutionConcentrationFinal 
         Height          =   315
         Left            =   3360
         Style           =   2  'Dropdown List
         TabIndex        =   44
         Tag             =   "7090"
         Top             =   1680
         Width           =   1600
      End
      Begin VB.TextBox txtTotalVolume 
         Height          =   315
         Left            =   2160
         TabIndex        =   49
         Tag             =   "11295"
         Text            =   "15"
         ToolTipText     =   "Volume of solvent the compound is dissolved in"
         Top             =   2640
         Width           =   1015
      End
      Begin VB.ComboBox cboTotalVolume 
         Height          =   315
         Left            =   3360
         Style           =   2  'Dropdown List
         TabIndex        =   50
         Tag             =   "11030"
         Top             =   2640
         Width           =   1600
      End
      Begin VB.TextBox txtDilutionConcentrationFinal 
         Height          =   315
         Left            =   2160
         TabIndex        =   43
         Tag             =   "11275"
         Text            =   "2"
         ToolTipText     =   "Molarity of compound in the solvent"
         Top             =   1680
         Width           =   1015
      End
      Begin VB.Label lblDilutingSolventVolume 
         Caption         =   "Volume of Solvent used for &Dilution"
         Height          =   405
         Left            =   120
         TabIndex        =   45
         Tag             =   "11280"
         Top             =   2120
         Width           =   1905
      End
      Begin VB.Label lblTotalVolume 
         Caption         =   "&Total Final Volume"
         Height          =   405
         Left            =   120
         TabIndex        =   48
         Tag             =   "11290"
         Top             =   2630
         Width           =   1905
      End
      Begin VB.Label lblDilutionConcentrationFinal 
         Caption         =   "&Final Concentration"
         Height          =   405
         Left            =   120
         TabIndex        =   42
         Tag             =   "11270"
         Top             =   1665
         Width           =   1905
      End
      Begin VB.Label lblStockSolutionVolume 
         Caption         =   "Volume of &Stock Solution"
         Height          =   405
         Left            =   120
         TabIndex        =   39
         Tag             =   "11260"
         Top             =   1185
         Width           =   1905
      End
      Begin VB.Label lblDilutionConcentrationInitial 
         Caption         =   "&Initial Concentration"
         Height          =   405
         Left            =   120
         TabIndex        =   36
         Tag             =   "11250"
         Top             =   705
         Width           =   1905
      End
   End
   Begin VB.ComboBox cboAction 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Tag             =   "11010"
      ToolTipText     =   "Perform conversions between different amounts of compound or perform molarity-related calculations"
      Top             =   1680
      Width           =   2955
   End
   Begin VB.Frame fraWeightSource 
      BorderStyle     =   0  'None
      Caption         =   "Weight Source"
      Height          =   1575
      Left            =   240
      TabIndex        =   52
      Top             =   0
      Width           =   5655
      Begin VB.TextBox txtMWTValue 
         Height          =   285
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   63
         Tag             =   "7600"
         Text            =   "100"
         ToolTipText     =   "Enter custom numerical mass for use in computations"
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox txtCustomMass 
         Height          =   285
         Left            =   3120
         TabIndex        =   55
         Tag             =   "7600"
         Text            =   "100"
         ToolTipText     =   "Enter custom numerical mass for use in computations"
         Top             =   1200
         Width           =   1575
      End
      Begin VB.OptionButton optWeightSource 
         Caption         =   "&Enter custom numerical mass"
         Height          =   255
         HelpContextID   =   4010
         Index           =   1
         Left            =   0
         TabIndex        =   54
         Tag             =   "7760"
         Top             =   240
         Width           =   5475
      End
      Begin VB.OptionButton optWeightSource 
         Caption         =   "&Use mass of compound in current formula"
         Height          =   255
         HelpContextID   =   4010
         Index           =   0
         Left            =   0
         TabIndex        =   53
         Tag             =   "7750"
         Top             =   0
         Value           =   -1  'True
         Width           =   5475
      End
      Begin RichTextLib.RichTextBox rtfCurrentFormula 
         Height          =   495
         Left            =   2040
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   600
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   873
         _Version        =   393217
         Enabled         =   0   'False
         MultiLine       =   0   'False
         Appearance      =   0
         TextRTF         =   $"MMCONV.frx":08CA
      End
      Begin VB.Label lblCustomMass 
         Caption         =   "Custom Mass:"
         Height          =   255
         Left            =   2040
         TabIndex        =   57
         Tag             =   "7480"
         Top             =   1230
         Width           =   2175
      End
      Begin VB.Label lblFormula 
         Caption         =   "Current Formula is"
         Height          =   255
         Left            =   120
         TabIndex        =   60
         Tag             =   "7460"
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label lblMWT 
         Caption         =   "MW ="
         Height          =   255
         Left            =   120
         TabIndex        =   59
         Tag             =   "7470"
         Top             =   1230
         Width           =   1335
      End
      Begin VB.Label lblCustomMassUnits 
         Caption         =   "g/mole"
         Height          =   255
         Left            =   4800
         TabIndex        =   58
         Tag             =   "7570"
         Top             =   1200
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "Cl&ose"
      Height          =   360
      Left            =   6720
      TabIndex        =   51
      Tag             =   "4000"
      Top             =   5520
      Width           =   1035
   End
   Begin VB.Label lblDensity 
      Caption         =   "g/mL"
      Height          =   255
      Left            =   4365
      TabIndex        =   62
      Tag             =   "11080"
      Top             =   1680
      Width           =   615
   End
End
Attribute VB_Name = "frmMMConvert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum mmaMoleMassActionConstants
    mmaConvertAmounts = 0
    mmaFindConcentration
    mmaDilutionCalculations
End Enum

' Constants for cboDilutionMode Combo Box on frmMMConvert
Private Enum dmDilutionModeConstants
     dmFindRequiredDilutionVolumes = 0
     dmFindRequiredTotalVolume = 1
     dmFindFinalMolarity = 2
     dmFindInitialMolarity = 3
End Enum

Private Const TextBoxIDMaxIndex = 10
Private Enum mmtTextBoxIDConstants
    mmtNone = 0
    mmtSampleMass = 1
    mmtSampleDensity = 2
    mmtQuantityAmount = 3
    mmtQuantityVolume = 4
    mmtQuantityConcentration = 5
    mmtDilutionInitialConcentration = 6
    mmtDilutionVolumeStock = 7
    mmtDilutionFinalConcentration = 8
    mmtDilutionVolumeDilutingSolvent = 9
    mmtDilutionTotalVolume = 10
End Enum

' Constants for cboMassError combo box
Private Enum meMassErrorConstants
    meMZ = 0
    mePPM = 1
End Enum

' Note: Other forms access this variable
Public eMassMode As mmcMassModeConstants

Private mDelayUpdateAll As Boolean                ' When true, the values will not auto-update

Private mDelayUpdatingDilutionVolumes As Boolean
Private mUpdatingDilutingSolventVolume As Boolean
Private mUpdatingTotalSolventVolume As Boolean
Private mUpdatingLinkedMolarities As Boolean

Private mLastControlChanged As mmtTextBoxIDConstants            ' Records the last box a key was pressed in
Private mVolumeIndexStart As Integer
Private mVolumeIndexEnd As Integer

Private objMMConverter As New MWMoleMassDilutionClass

Private Sub ComputeMassRanges()
    
    ComputeMassRangesWork txtMH, txtMZLow, txtMZHigh
    
    ComputeMassRangesWork txtMHAlt, txtMZLowAlt, txtMZHighAlt
    
End Sub

Private Sub ComputeMassRangesWork(ByRef objMH As TextBox, ByRef objMZLow As TextBox, ByRef objMZHigh As TextBox)

    Dim dblMassError As Double
    Dim dblMassErrorMZ As Double
    Dim dblCurrentMZ As Double

    dblMassError = CDblSafe(txtMassError)
    dblCurrentMZ = CDblSafe(objMH.Text)
    
    If cboMassError.ListIndex = meMassErrorConstants.mePPM Then
        dblMassErrorMZ = PPMtoMass(dblMassError, dblCurrentMZ)
    Else
        dblMassErrorMZ = dblMassError
    End If
    
    objMZLow.Text = Round(objMH - dblMassErrorMZ, 6)
    objMZHigh.Text = Round(objMH + dblMassErrorMZ, 6)

End Sub

Private Sub ComputeTargetConcentration()
    Dim dblFromNum As Double
    Dim dblToNum As Double
    
    dblFromNum = CDblSafe(txtFromNum.Text)
    dblToNum = objMMConverter.ConvertAmount(dblFromNum, cboFrom.ListIndex, cboTo.ListIndex)
    
    If dblToNum < 0 Then
        lblToNum.Caption = ""
    Else
        FormatLabel lblToNum, dblToNum
    End If

End Sub

Private Sub ConvertSequenceMH(blnFavorMH As Boolean)
    ' When blnFavorMH = True then computes MHAlt using MH
    ' Otherwise, computes MH using MHAlt
    
    Static blnWorking As Boolean
    Dim dblMH As Double, dblMHAlt As Double
    Dim intChargePrimary As Integer, intChargeAlt As Integer
    
    If mDelayUpdateAll Then Exit Sub

    If blnWorking Then Exit Sub
    blnWorking = True
    
    intChargePrimary = cboMH.ListIndex
    If intChargePrimary < 0 Then intChargePrimary = 0
    
    intChargeAlt = cboMHAlternate.ListIndex
    If intChargeAlt < 0 Then intChargeAlt = 0
    
    If blnFavorMH Then
        dblMH = CDblSafe(txtMH)
        If dblMH > 0 Then
            dblMHAlt = objMwtWin.ConvoluteMass(dblMH, intChargePrimary, intChargeAlt)
        Else
            dblMHAlt = 0
        End If
        txtMHAlt = Round(dblMHAlt, 6)
    Else
        dblMHAlt = CDblSafe(txtMHAlt)
        If dblMHAlt > 0 Then
            dblMH = objMwtWin.ConvoluteMass(dblMHAlt, intChargeAlt, intChargePrimary)
        Else
            dblMH = 0
        End If
        txtMH = Round(dblMH, 6)
    End If
    
    ComputeMassRanges
    
    blnWorking = False
    
End Sub

Private Sub FindDilutionValues(Optional blnUpdateDilutingSolventVolume As Boolean = True)
    
    Dim dblVolumeStock As Double
    Dim dblDilutingSolventVolume As Double
    Dim dblTotalVolume As Double
    Dim dblConcentration As Double
    
    Static intCallLevel As Integer
    
    ' The intCallLevel variable and the following assertion are here to make sure
    '  we don't accidentally get into redundant loop with calling this Sub
    intCallLevel = intCallLevel + 1
    Debug.Assert intCallLevel = 1
    
    ' Find the desired item
    ' After finding the item, need to update the variable for the item just found
    Select Case cboDilutionMode.ListIndex
    Case dmFindRequiredDilutionVolumes
        dblVolumeStock = objMMConverter.ComputeDilutionRequiredStockAndDilutingSolventVolumes(cboStockSolutionVolume.ListIndex, dblDilutingSolventVolume, cboDilutingSolventVolume.ListIndex)
        
        If dblDilutingSolventVolume < 0 Then
            txtStockSolutionVolume = ""
            txtDilutingSolventVolume = ""
        Else
            FormatTextBox txtStockSolutionVolume, dblVolumeStock
            FormatTextBox txtDilutingSolventVolume, dblDilutingSolventVolume
        End If
        
    Case dmFindFinalMolarity
        dblConcentration = objMMConverter.ComputeDilutionFinalConcentration(cboDilutionConcentrationFinal.ListIndex)
        FormatTextBox txtDilutionConcentrationFinal, dblConcentration
        If blnUpdateDilutingSolventVolume Then UpdateDilutingSolventVolume
        
    Case dmFindInitialMolarity
        dblConcentration = objMMConverter.ComputeDilutionInitialConcentration(cboDilutionConcentrationInitial.ListIndex)
        FormatTextBox txtDilutionConcentrationInitial, dblConcentration
        If blnUpdateDilutingSolventVolume Then UpdateDilutingSolventVolume

    Case Else   ' Includes dmFindRequiredTotalVolume
        dblTotalVolume = objMMConverter.ComputeDilutionTotalVolume(cboTotalVolume.ListIndex, dblDilutingSolventVolume, cboDilutingSolventVolume.ListIndex)
        If dblDilutingSolventVolume <= 0 Then
            txtDilutingSolventVolume = ""
        Else
            FormatTextBox txtDilutingSolventVolume, dblDilutingSolventVolume
        End If
        FormatTextBox txtTotalVolume, dblTotalVolume
    End Select

    UpdateDynamicDilutionFrameLabel
    
    intCallLevel = intCallLevel - 1
    mLastControlChanged = mmtNone
End Sub

Private Sub FindQuantityAmount()
    Dim dblAmount As Double
    
    dblAmount = objMMConverter.ComputeQuantityAmount(cboFrom.ListIndex)
    FormatTextBox txtFromNum, dblAmount

    mLastControlChanged = mmtNone
End Sub

Private Sub FindQuantityConcentration()
    Dim dblConcentration As Double
    
    dblConcentration = objMMConverter.ComputeQuantityConcentration(cboConcentration.ListIndex)
    FormatTextBox txtConcentration, dblConcentration
    
    mLastControlChanged = mmtNone
End Sub

Private Sub FindQuantityVolume()
    Dim dblVolume As Double
    
    dblVolume = objMMConverter.ComputeQuantityVolume(cboVolume.ListIndex)
    FormatTextBox txtVolume, dblVolume
    
    mLastControlChanged = mmtNone

End Sub

Private Function GetWorkingMass() As Double
    If eMassMode = mmcComputedMass Then
        GetWorkingMass = CDblSafe(txtMWTValue.Text)
    Else
        GetWorkingMass = CDblSafe(txtCustomMass.Text)
    End If
    
End Function

Private Sub HandleDilutingSolventVolumeChange()
    
    If cboDilutionMode.ListIndex = dmFindRequiredDilutionVolumes Or _
       cboDilutionMode.ListIndex = dmFindRequiredTotalVolume Then
        Exit Sub
    End If
    
    SynchronizeQuantitiesWithDLL mmtDilutionVolumeDilutingSolvent
    
    If mUpdatingDilutingSolventVolume Then Exit Sub
    
    UpdateTotalSolventVolume
    
    FindDilutionValues False

End Sub

Private Sub HighlightTargetBoxes()
    ' Change background color of text boxes to indicate which is being found
    
    Dim ctlThisControl As Control, strControlType As String
    Dim lngHighlightedColor As Long
    
    For Each ctlThisControl In Me.Controls
        strControlType = TypeName(ctlThisControl)
        If strControlType = "TextBox" Then
            ctlThisControl.BackColor = QBColor(COLOR_WHITE)
        End If
    Next

    ' 7, 10, and 14 are good colors for QBColor()
    lngHighlightedColor = QBColor(COLOR_COMPUTEDQUANTITY)
    Select Case cboDilutionMode.ListIndex
    Case dmFindRequiredTotalVolume                      ' Find Required Total Volume
        txtDilutingSolventVolume.BackColor = lngHighlightedColor
        txtTotalVolume.BackColor = lngHighlightedColor
    Case dmFindFinalMolarity                      ' Find Final Molarity
        txtDilutionConcentrationFinal.BackColor = lngHighlightedColor
    Case dmFindInitialMolarity                      ' Find Initial Molarity
        txtDilutionConcentrationInitial.BackColor = lngHighlightedColor
    Case Else
        ' Includes case 0       ' Find Required Dilution Volumes
        txtStockSolutionVolume.BackColor = lngHighlightedColor
        txtDilutingSolventVolume.BackColor = lngHighlightedColor
    End Select
    
End Sub

Private Sub UpdateDilutingSolventVolume()
    Dim dblDilutingSolventVolume As Double
    
    mUpdatingDilutingSolventVolume = True
    
    dblDilutingSolventVolume = objMMConverter.GetDilutionVolumeDilutingSolvent(cboDilutingSolventVolume.ListIndex)
    
    If dblDilutingSolventVolume < 0 Then
        txtDilutingSolventVolume = ""
    ElseIf dblDilutingSolventVolume = 0 Then
        txtDilutingSolventVolume = "0"
    Else
        FormatTextBox txtDilutingSolventVolume, dblDilutingSolventVolume
    End If
    
    mUpdatingDilutingSolventVolume = False
    
End Sub

Private Sub UpdateTotalSolventVolume()
    Dim dblStockSolutionVolume As Double
    Dim dblDilutingSolventVolume As Double
    Dim dblFinalTotalVolume As Double

    mUpdatingTotalSolventVolume = True
    
    With objMMConverter
        dblStockSolutionVolume = .GetDilutionVolumeStockSolution(uevML)
        dblDilutingSolventVolume = .GetDilutionVolumeDilutingSolvent(uevML)
            
        dblFinalTotalVolume = .ConvertVolumeExtended(dblStockSolutionVolume + dblDilutingSolventVolume, uevML, cboTotalVolume.ListIndex)
        
        If dblFinalTotalVolume < 0 Then
            txtTotalVolume.Text = ""
        Else
            FormatTextBox txtTotalVolume, dblFinalTotalVolume
        End If
    End With
    
    mUpdatingTotalSolventVolume = False
    
End Sub

Private Sub PopulateComboBoxes()
    Dim strAmountString As String, strVolumeString As String, strMolarString As String
    Dim intIndex As Integer
    
    ' Load Amount types in the Combo Boxes
    
    SetDelayUpdate True
    
    PopulateComboBox cboAction, True, "Convert Amounts|Find Concentration|Dilution Calculations", 1
    PopulateComboBox cboDilutionMode, True, "Find Required Dilution Volumes|Find Required Total Volume|Find Final Concentration|Find Initial Concentration", 0
    
    strAmountString = "Moles|Millimoles|microMoles|nanoMoles|picoMoles|femtoMoles|attoMoles|Kilograms|Grams|Milligrams|Micrograms|Pounds|Ounces|Liters|Deciliters|Milliliters|Microliters|Nanoliters|Picoliters|Gallons|Quarts|Pints"
    PopulateComboBox cboFrom, True, strAmountString, uamGrams
    PopulateComboBox cboTo, True, strAmountString, uamMoles
    
    Debug.Assert cboFrom.ListCount = Val(objMMConverter.AmountsUnitListCount)
    
    strVolumeString = "Liters|Deciliters|Milliliters|Microliters|Nanoliters|Picoliters|Gallons|Quarts|Pints"
    PopulateComboBox cboVolume, True, strVolumeString, uevML
    PopulateComboBox cboStockSolutionVolume, True, strVolumeString, uevML
    PopulateComboBox cboDilutingSolventVolume, True, strVolumeString, uevML
    PopulateComboBox cboTotalVolume, True, strVolumeString, uevML
   
    strMolarString = "Molar|milliMolar|microMolar|nanoMolar|picoMolar|femtoMolar|attoMolar|mg/dL|mg/mL|ug/mL|ng/mL|ug/uL|ng/uL"
    PopulateComboBox cboConcentration, True, strMolarString, ummcMolar
    PopulateComboBox cboDilutionConcentrationInitial, True, strMolarString, ummcMolar
    PopulateComboBox cboDilutionConcentrationFinal, True, strMolarString, ummcMolar
    
    With cboMassError
        .Clear
        .AddItem "m/z"
        .AddItem "ppm"
        .ListIndex = 1      ' ppm
    End With
    
    With cboMH
        .Clear
        .AddItem "M (uncharged)"
        .AddItem "[M+H]1+"
        For intIndex = 2 To 9
            .AddItem "[M+" & Trim(intIndex) & "H]" & Trim(intIndex) & "+"
        Next intIndex
        .ListIndex = 0    ' Uncharged
    End With
    
    
    With cboMHAlternate
        .Clear
        .AddItem "M (uncharged)"
        .AddItem "[M+H]1+"
        For intIndex = 2 To 9
            .AddItem "[M+" & Trim(intIndex) & "H]" & Trim(intIndex) & "+"
        Next intIndex
        .ListIndex = 1    ' Charge of 1+
    End With
    
    SetDelayUpdate False
End Sub

Private Sub PositionFormControls()
    
    fraWeightSource.Top = 100
    fraWeightSource.Left = 240
    PositionWeightSourceframeControls Me
       
    cboAction.Top = 1680
    cboAction.Left = lblFormula.Left
    
    txtDensity.Top = cboAction.Top + 20
    txtDensity.Left = 3240
    lblDensity.Top = txtDensity.Top
    lblDensity.Left = 4365
    
    fraConvertAmounts.Top = 2040
    fraConvertAmounts.Left = 120
    
    txtFromNum.Top = 300
    txtFromNum.Left = 120
    cboFrom.Top = txtFromNum.Top
    cboFrom.Left = 1320
    cmdFindAmount.Top = txtFromNum.Top - 40
    cmdFindAmount.Left = 3075
    
    lblToNum.Top = 820
    lblToNum.Left = txtFromNum.Left + 40
    cboTo.Top = lblToNum.Top - 40
    cboTo.Left = cboFrom.Left
    
    txtVolume.Top = 900
    txtVolume.Left = txtFromNum.Left
    cboVolume.Top = txtVolume.Top
    cboVolume.Left = cboFrom.Left
    cmdFindVolume.Left = cmdFindAmount.Left
    cmdFindVolume.Top = txtVolume.Top - 40
    txtConcentration.Top = 1500
    txtConcentration.Left = txtFromNum.Left
    cboConcentration.Top = txtConcentration.Top
    cboConcentration.Left = cboFrom.Left
    cmdFindConcentration.Left = cmdFindAmount.Left
    cmdFindConcentration.Top = txtConcentration.Top - 40

    fraDilutionCalcs.Top = 1560
    fraDilutionCalcs.Left = 5160
    
    txtDilutionConcentrationInitial.Top = 720
    txtDilutionConcentrationInitial.Left = 2160
    lblDilutionConcentrationInitial.Top = txtDilutionConcentrationInitial.Top - 15
    lblDilutionConcentrationInitial.Left = 120
    cboDilutionConcentrationInitial.Top = txtDilutionConcentrationInitial.Top
    cboDilutionConcentrationInitial.Left = 3360
    
    txtStockSolutionVolume.Top = 1200
    txtStockSolutionVolume.Left = txtDilutionConcentrationInitial.Left
    lblStockSolutionVolume.Top = txtStockSolutionVolume.Top - 15
    lblStockSolutionVolume.Left = lblDilutionConcentrationInitial.Left
    cboStockSolutionVolume.Top = txtStockSolutionVolume.Top
    cboStockSolutionVolume.Left = cboDilutionConcentrationInitial.Left
   
    txtDilutionConcentrationFinal.Top = 1680
    txtDilutionConcentrationFinal.Left = txtDilutionConcentrationInitial.Left
    lblDilutionConcentrationFinal.Top = txtDilutionConcentrationFinal.Top - 15
    lblDilutionConcentrationFinal.Left = lblDilutionConcentrationInitial.Left
    cboDilutionConcentrationFinal.Top = txtDilutionConcentrationFinal.Top
    cboDilutionConcentrationFinal.Left = cboDilutionConcentrationInitial.Left
   
    txtDilutingSolventVolume.Top = 2160
    txtDilutingSolventVolume.Left = txtDilutionConcentrationInitial.Left
    lblDilutingSolventVolume.Top = txtDilutingSolventVolume.Top - 15
    lblDilutingSolventVolume.Left = lblDilutionConcentrationInitial.Left
    cboDilutingSolventVolume.Top = txtDilutingSolventVolume.Top
    cboDilutingSolventVolume.Left = cboDilutionConcentrationInitial.Left
   
    txtTotalVolume.Top = 2640
    txtTotalVolume.Left = txtDilutionConcentrationInitial.Left
    lblTotalVolume.Top = txtTotalVolume.Top - 15
    lblTotalVolume.Left = lblDilutionConcentrationInitial.Left
    cboTotalVolume.Top = txtTotalVolume.Top
    cboTotalVolume.Left = cboDilutionConcentrationInitial.Left

    ShowHideFormControls True
    
End Sub

Private Function PPMtoMass(ByVal dblPPMToConvert As Double, ByVal dblCurrentMZ As Double) As Double
    ' Converts dblMassToConvert to ppm, based on the value of dblCurrentMZ
    
    If dblCurrentMZ > 0 Then
        PPMtoMass = dblPPMToConvert * dblCurrentMZ / 1000000#
    Else
        PPMtoMass = 0
    End If
End Function

Private Sub RemoveNegativeValue(txtThisTextBox As TextBox)
    If IsNumeric(txtThisTextBox.Text) Then
        If CDbl(txtThisTextBox) < 0 Then
            txtThisTextBox = 0
        End If
    Else
        txtThisTextBox.Text = ""
    End If
End Sub

Public Sub SetDelayUpdate(blnTurnOnDelay As Boolean)
    mDelayUpdateAll = blnTurnOnDelay
    
    If Not mDelayUpdateAll Then
        ' Just turned off delaying
        ' Make sure the values are correct
        
        SynchronizeAllQuantitiesWithDLL
        
        FindQuantityConcentration
        
        HighlightTargetBoxes
        FindDilutionValues
        
        ConvertSequenceMH True
    End If
    
End Sub

Public Sub SetDynamicToolTips()
    Select Case cboAction.ListIndex
    Case mmaConvertAmounts
        txtFromNum.ToolTipText = LookupToolTipLanguageCaption(11100, "Amount of compound to convert from")
        cboFrom.ToolTipText = LookupToolTipLanguageCaption(11020, "Units of amount to convert from")
    Case Else
        txtFromNum.ToolTipText = LookupToolTipLanguageCaption(11105, "Amount of compound to be dissolved in solvent")
        cboFrom.ToolTipText = LookupToolTipLanguageCaption(11025, "Units of amount to use for concentration calculation")
    End Select
End Sub

Private Sub ShowHideDensity()
    Dim blnShowDensity As Boolean
    
    If (cboFrom.ListIndex >= mVolumeIndexStart And cboFrom.ListIndex <= mVolumeIndexEnd) Then
        blnShowDensity = True
    ElseIf cboAction.ListIndex = mmaConvertAmounts And (cboTo.ListIndex >= mVolumeIndexStart And cboTo.ListIndex <= mVolumeIndexEnd) Then
        blnShowDensity = True
    Else
        blnShowDensity = False
    End If
    
    txtDensity.Visible = blnShowDensity
    lblDensity.Visible = blnShowDensity
    
End Sub

Private Sub ShowHideFormControls(Optional blnForceResize As Boolean = True)
    Dim blnVolAndMolarVisible As Boolean, blnDilutionFrameVisible As Boolean
    Dim intFindAmountsFrameHeight As Integer, intFindAmountsFrameWidth As Integer
    Dim lngPreferredWidth As Long, lngPreferredHeight As Long
    
    If cboAction.ListIndex = mmaConvertAmounts Then
        ' Set Molarity related objects to invisible
        blnVolAndMolarVisible = False
        blnDilutionFrameVisible = False
        intFindAmountsFrameHeight = 1300
        intFindAmountsFrameWidth = 3200
    Else
        If cboAction.ListIndex = mmaDilutionCalculations Then
            blnDilutionFrameVisible = True
        Else
            blnDilutionFrameVisible = False
        End If
        
        blnVolAndMolarVisible = True
        intFindAmountsFrameHeight = 2120
        intFindAmountsFrameWidth = 4935
    End If
    
    fraConvertAmounts.Height = intFindAmountsFrameHeight
    fraConvertAmounts.Width = intFindAmountsFrameWidth
        
    lblToNum.Visible = Not blnVolAndMolarVisible
    cboTo.Visible = Not blnVolAndMolarVisible
    
    cmdFindAmount.Visible = blnVolAndMolarVisible
    txtVolume.Visible = blnVolAndMolarVisible
    cboVolume.Visible = blnVolAndMolarVisible
    cmdFindVolume.Visible = blnVolAndMolarVisible
    
    txtConcentration.Visible = blnVolAndMolarVisible
    cboConcentration.Visible = blnVolAndMolarVisible
    cmdFindConcentration.Visible = blnVolAndMolarVisible
    
    fraDilutionCalcs.Visible = blnDilutionFrameVisible
    chkLinkMolarities.Visible = blnDilutionFrameVisible
    chkLinkDilutionVolumeUnits.Visible = blnDilutionFrameVisible
    
    If blnDilutionFrameVisible Then
        fraMZConversion.Top = fraDilutionCalcs.Top + fraDilutionCalcs.Height + 100
        
        With fraMZConversion
            cmdOK.Top = .Top + .Height / 2 - (cmdOK.Height / 2)
            cmdOK.Left = .Left + .Width + 150
        End With
    Else
        fraMZConversion.Top = fraConvertAmounts.Top + fraConvertAmounts.Height + 100
    
        cmdOK.Top = fraMZConversion.Top - cmdOK.Height - 100
        cmdOK.Left = fraMZConversion.Left + fraMZConversion.Width - cmdOK.Width
    End If
    
    
    If cChkBox(frmProgramPreferences.chkShowToolTips) Then SetDynamicToolTips
    
    If Me.WindowState <> vbMinimized Then
        If blnVolAndMolarVisible Then
            If blnDilutionFrameVisible Then
                lngPreferredHeight = 7300
            Else
                lngPreferredHeight = 6750
            End If
        Else
            lngPreferredHeight = 5900
        End If
            
        If blnDilutionFrameVisible Then
            lngPreferredWidth = 10550
        Else
            lngPreferredWidth = 6700
        End If
    Else
        ' Cannot change form height or width when minimized
        ' Does not generate an error in the VB IDE but does once compiled to a .Exe
    End If
    
    If Me.WindowState = vbNormal Then
        If Me.Height > lngPreferredHeight Or blnForceResize Then
            Me.Height = lngPreferredHeight
        End If
        
        If Me.Width > lngPreferredWidth Or blnForceResize Then
            Me.Width = lngPreferredWidth
        End If
    End If
End Sub

Private Sub ShowHideMassInputControls()
    Dim blnCustomVisible As Boolean
    
    If optWeightSource(0).value = True Then
        eMassMode = mmcComputedMass
    Else
        eMassMode = mmcCustomMass
    End If
    
    If eMassMode = mmcComputedMass Then
        blnCustomVisible = False
    Else
        blnCustomVisible = True
    End If
    
    lblFormula.Visible = Not blnCustomVisible
    rtfCurrentFormula.Visible = Not blnCustomVisible
    lblMWT.Visible = Not blnCustomVisible
    txtMWTValue.Visible = Not blnCustomVisible
    
    lblCustomMass.Visible = blnCustomVisible
    txtCustomMass.Visible = blnCustomVisible
End Sub

Private Sub SynchronizeAllQuantitiesWithDLL()
    Dim intIndex As Integer
    
    For intIndex = 0 To TextBoxIDMaxIndex
        SynchronizeQuantitiesWithDLL CInt(intIndex)
    Next intIndex
End Sub

Private Sub SynchronizeDilutionVolumes(cboThisComboBox As ComboBox)
    
    ' The following is necessary because when the form is loading, the cboStockSolutionVolume combobox
    '  gets populated and has its .ListIndex value set, but the other two volume comboboxes have
    '  no items yet, so an error occurs when trying to set their .ListIndex value
    On Error Resume Next
    
    If cChkBox(chkLinkDilutionVolumeUnits) Then
        mDelayUpdatingDilutionVolumes = True
        With cboThisComboBox
            ' Note:  The order of updating the ComboBoxes is important -- do not change
            If cboStockSolutionVolume.Name <> .Name Then cboStockSolutionVolume.ListIndex = .ListIndex
            If cboTotalVolume.Name <> .Name Then cboTotalVolume.ListIndex = .ListIndex
            If cboDilutingSolventVolume.Name <> .Name Then cboDilutingSolventVolume.ListIndex = .ListIndex
        End With
        mDelayUpdatingDilutionVolumes = False
        
        SynchronizeQuantitiesWithDLL mmtDilutionVolumeStock
        SynchronizeQuantitiesWithDLL mmtDilutionVolumeDilutingSolvent
        SynchronizeQuantitiesWithDLL mmtDilutionTotalVolume

        FindDilutionValues
    End If
End Sub

Private Sub SynchronizeLinkedMolarities(blnFavorDilutionMolarity As Boolean)

    If mUpdatingLinkedMolarities Then Exit Sub
    
    If cChkBox(chkLinkMolarities) Then
        mUpdatingLinkedMolarities = True
        
        If blnFavorDilutionMolarity Then
            If IsNumeric(txtDilutionConcentrationInitial.Text) Then
                txtConcentration.Text = txtDilutionConcentrationInitial.Text
                mLastControlChanged = mmtQuantityConcentration
                cboConcentration.ListIndex = cboDilutionConcentrationInitial.ListIndex
                mLastControlChanged = mmtQuantityConcentration
            End If
        Else
            If IsNumeric(txtConcentration.Text) And cboDilutionMode.ListIndex <> dmFindInitialMolarity Then
                txtDilutionConcentrationInitial.Text = txtConcentration.Text
                mLastControlChanged = mmtQuantityConcentration
                cboDilutionConcentrationInitial.ListIndex = cboConcentration.ListIndex
                mLastControlChanged = mmtQuantityConcentration
            End If
        End If
        
        mUpdatingLinkedMolarities = False
    End If

End Sub

Private Sub SynchronizeQuantitiesWithDLL(eTextBoxID As mmtTextBoxIDConstants)
    
    If mDelayUpdateAll Then Exit Sub
    
    With objMMConverter
        Select Case eTextBoxID
        Case mmtSampleMass
            .SetSampleMass GetWorkingMass()
            
        Case mmtSampleDensity
            .SetSampleDensity CDblSafe(txtDensity.Text)
        
        Case mmtQuantityAmount
            .SetQuantityAmount CDblSafe(txtFromNum.Text), cboFrom.ListIndex
            
        Case mmtQuantityVolume
            .SetQuantityVolume CDblSafe(txtVolume.Text), cboVolume.ListIndex
        
        Case mmtQuantityConcentration
            .SetQuantityConcentration CDblSafe(txtConcentration.Text), cboConcentration.ListIndex
        
        Case mmtDilutionInitialConcentration
            .SetDilutionInitialConcentration CDblSafe(txtDilutionConcentrationInitial.Text), cboDilutionConcentrationInitial.ListIndex
        
        Case mmtDilutionVolumeStock
            .SetDilutionVolumeStockSolution CDblSafe(txtStockSolutionVolume.Text), cboStockSolutionVolume.ListIndex
        
        Case mmtDilutionFinalConcentration
            .SetDilutionFinalConcentration CDblSafe(txtDilutionConcentrationFinal.Text), cboDilutionConcentrationFinal.ListIndex
        
        Case mmtDilutionVolumeDilutingSolvent
            .SetDilutionVolumeDilutingSolvent CDblSafe(txtDilutingSolventVolume.Text), cboDilutingSolventVolume.ListIndex
        
        Case mmtDilutionTotalVolume
            .SetDilutionTotalFinalVolume CDblSafe(txtTotalVolume.Text), cboTotalVolume.ListIndex
        
        Case Else
            If eTextBoxID <> 0 Then
                ' This shouldn't happen
                Debug.Assert False
            End If
        End Select
    End With
    
End Sub

Private Sub UpdateDynamicDilutionFrameLabel()
    Dim blnDilutionMode As Boolean
    
    blnDilutionMode = True
    If IsNumeric(txtDilutionConcentrationInitial) And IsNumeric(txtDilutionConcentrationFinal) Then
        If CDblSafe(txtDilutionConcentrationInitial) < CDblSafe(txtDilutionConcentrationFinal) Then
            blnDilutionMode = False
        End If
    End If
    
    If blnDilutionMode Then
        fraDilutionCalcs.Caption = LookupLanguageCaption(11200, "Dilution Calculations")
    Else
        fraDilutionCalcs.Caption = LookupLanguageCaption(11205, "Evaporation or Sublimation Calculations")
    End If
End Sub

Private Sub cboAction_Click()
    ShowHideFormControls True
    ShowHideDensity
    ComputeTargetConcentration
End Sub

Private Sub cboDilutionMode_Click()
    HighlightTargetBoxes
    FindDilutionValues
End Sub

Private Sub cboFrom_Click()
    
    ShowHideDensity
    
    If cboAction.ListIndex > mmaConvertAmounts And mLastControlChanged <> mmtQuantityAmount Then
        If Not mDelayUpdateAll Then FindQuantityAmount
    Else
        SynchronizeQuantitiesWithDLL mmtQuantityAmount
    End If

    ComputeTargetConcentration

End Sub

Private Sub cboConcentration_Click()
    If mLastControlChanged <> mmtQuantityConcentration And Not cChkBox(chkLinkMolarities.value) Then
        ' Auto-find the concentration for the new units, but only if user did not just
        ' enter a custom value and if not linking molarities
        If Not mDelayUpdateAll Then FindQuantityConcentration
    Else
        SynchronizeQuantitiesWithDLL mmtQuantityConcentration
    End If
    
    SynchronizeLinkedMolarities False
End Sub

Private Sub cboDilutionConcentrationFinal_Click()
    If mDelayUpdateAll Then Exit Sub
    
    SynchronizeQuantitiesWithDLL mmtDilutionFinalConcentration
    
    FindDilutionValues
End Sub

Private Sub cboDilutionConcentrationInitial_Click()
    If mDelayUpdateAll Then Exit Sub
    
    SynchronizeQuantitiesWithDLL mmtDilutionInitialConcentration
    
    FindDilutionValues
    SynchronizeLinkedMolarities True
End Sub

Private Sub cboMassError_Click()
  ComputeMassRanges
End Sub

Private Sub cboMH_Click()
    ConvertSequenceMH False
End Sub

Private Sub cboMHAlternate_Click()
    ConvertSequenceMH True
End Sub

Private Sub cboTo_Click()
    
    ShowHideDensity
    ComputeTargetConcentration

End Sub

Private Sub cboVolume_Click()
    If mLastControlChanged <> mmtQuantityVolume Then
        If Not mDelayUpdateAll Then FindQuantityVolume
    Else
        SynchronizeQuantitiesWithDLL mmtQuantityVolume
    End If
End Sub

Private Sub cboDilutingSolventVolume_Click()
    If mDelayUpdateAll Or mDelayUpdatingDilutionVolumes Then Exit Sub
    
    SynchronizeDilutionVolumes cboDilutingSolventVolume
    SynchronizeQuantitiesWithDLL mmtDilutionVolumeDilutingSolvent
    
    If IsNumeric(txtDilutingSolventVolume) Then
        HandleDilutingSolventVolumeChange
    End If

End Sub

Private Sub cboStockSolutionVolume_Click()
    If mDelayUpdateAll Or mDelayUpdatingDilutionVolumes Then Exit Sub
    
    SynchronizeDilutionVolumes cboStockSolutionVolume
    SynchronizeQuantitiesWithDLL mmtDilutionVolumeStock
    
    FindDilutionValues

End Sub

Private Sub cboTotalVolume_Click()
    If mDelayUpdateAll Or mDelayUpdatingDilutionVolumes Then Exit Sub
    
    SynchronizeDilutionVolumes cboTotalVolume
    SynchronizeQuantitiesWithDLL mmtDilutionTotalVolume
    
    FindDilutionValues
    
End Sub

Private Sub chkLinkDilutionVolumeUnits_Click()
    cboStockSolutionVolume_Click
End Sub

Private Sub chkLinkMolarities_Click()
    SynchronizeLinkedMolarities False
End Sub

Private Sub cmdFindAmount_Click()
    FindQuantityAmount
End Sub

Private Sub cmdFindConcentration_Click()
    FindQuantityConcentration
End Sub

Private Sub cmdFindVolume_Click()
    FindQuantityVolume
End Sub

Private Sub cmdOK_Click()
    HideFormShowMain Me
End Sub

Private Sub Form_Activate()
    ' Put window in center of screen (upper fourth vertically)
    SizeAndCenterWindow Me, cWindowExactCenter, 10700, 7300, True
    
    DisplayCurrentFormulaOnSubForm Me
    
    ' Make sure the values are sync'd up with the Dll
    SynchronizeAllQuantitiesWithDLL
    
    PossiblyHideMainWindow
    
    ShowHideMassInputControls
    ShowHideFormControls False
    
    Select Case objMwtWin.GetElementMode
    Case emAverageMass
        optElementMode(0).value = True
    Case emIsotopicMass
        optElementMode(1).value = True
    Case Else
        objMwtWin.SetElementMode emIsotopicMass
        optElementMode(1).value = True
    End Select
    
    ComputeTargetConcentration
    mLastControlChanged = mmtNone
    
    'ConvertSequenceMH
End Sub

Private Sub Form_Load()
    
    ' Turn off auto-compute
    objMMConverter.SetAutoComputeDilutionEnabled False
    objMMConverter.SetAutoComputeQuantityEnabled False

    mVolumeIndexStart = objMMConverter.AmountsUnitListVolumeIndexStart
    mVolumeIndexEnd = objMMConverter.AmountsUnitListVolumeIndexEnd
    
    PositionFormControls
    
    PopulateComboBoxes
            
    ShowHideDensity
        
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    QueryUnloadFormHandler Me, Cancel, UnloadMode
End Sub

Private Sub Form_Resize()
    ShowHideFormControls False
End Sub

Private Sub lblCustomMass_DblClick()
    ' This is sort of an Easter Egg function that will show frmStrings if
    '  lblCustomMass is double-clicked 3 times in a row, with a maximum of
    '  two seconds betwen clicks
    
    Static intDblClickCount As Integer
    Static lngLastClickTime As Long
    
    If Timer() < lngLastClickTime + 2 Then
        intDblClickCount = intDblClickCount + 1
        If intDblClickCount > 2 Then
            frmStrings.Show
            intDblClickCount = 0
        End If
    Else
        intDblClickCount = 1
    End If
    lngLastClickTime = Timer()
End Sub

Private Sub optElementMode_Click(Index As Integer)
    ' Elementweightmode = 1 means average weights, 2 means isotopic, and 3 means integer
    Dim eNewWeightMode As emElementModeConstants
    
    eNewWeightMode = Index + 1
        
    If eNewWeightMode <> objMwtWin.GetElementMode Then
        SwitchWeightMode eNewWeightMode, True, False
      
        ConvertSequenceMH True
    End If
    
End Sub

Private Sub optWeightSource_Click(Index As Integer)
    ShowHideMassInputControlsGlobal Me
    
    SynchronizeQuantitiesWithDLL mmtSampleMass
    SynchronizeQuantitiesWithDLL mmtQuantityAmount
    
    ComputeTargetConcentration
End Sub

Private Sub txtCustomMass_Change()
    SynchronizeQuantitiesWithDLL mmtSampleMass
    SynchronizeQuantitiesWithDLL mmtQuantityAmount
    SynchronizeQuantitiesWithDLL mmtDilutionInitialConcentration
    
    If cboDilutionMode.ListIndex <> dmFindFinalMolarity Then
        SynchronizeQuantitiesWithDLL mmtDilutionFinalConcentration
    End If
    
    ComputeTargetConcentration
    FindDilutionValues
    
    mLastControlChanged = mmtNone
End Sub

Private Sub txtCustomMass_GotFocus()
    HighlightOnFocus txtCustomMass
End Sub

Private Sub txtCustomMass_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtCustomMass, KeyAscii, True, True, True, False, True, False, False, False, False, True
End Sub

Private Sub txtDensity_Change()
    SynchronizeQuantitiesWithDLL mmtSampleDensity
    
    ComputeTargetConcentration
    mLastControlChanged = mmtQuantityVolume
End Sub

Private Sub txtDensity_GotFocus()
    HighlightOnFocus txtDensity
End Sub

Private Sub txtDensity_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtDensity, KeyAscii, True, True, True, False, True, False, False, False, False, True
End Sub

Private Sub txtFromNum_Change()
    SynchronizeQuantitiesWithDLL mmtQuantityAmount
    
    ComputeTargetConcentration
    mLastControlChanged = mmtQuantityAmount
End Sub

Private Sub txtFromNum_GotFocus()
    HighlightOnFocus txtFromNum
End Sub

Private Sub txtFromNum_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtFromNum, KeyAscii, True, True, True, False, True, False, False, False, False, True
End Sub

Private Sub txtConcentration_Change()
    SynchronizeQuantitiesWithDLL mmtQuantityConcentration
    mLastControlChanged = mmtQuantityConcentration
    
    SynchronizeLinkedMolarities False
End Sub

Private Sub txtConcentration_GotFocus()
    HighlightOnFocus txtConcentration
End Sub

Private Sub txtConcentration_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtConcentration, KeyAscii, True, True, True, False, True, False, False, False, False, True
End Sub

Private Sub txtDilutionConcentrationFinal_Change()
    If cboDilutionMode.ListIndex = dmFindFinalMolarity Then Exit Sub
    
    SynchronizeQuantitiesWithDLL mmtDilutionFinalConcentration
    FindDilutionValues
End Sub

Private Sub txtDilutionConcentrationFinal_GotFocus()
    HighlightOnFocus txtDilutionConcentrationFinal
End Sub

Private Sub txtDilutionConcentrationFinal_KeyPress(KeyAscii As Integer)
    If cboDilutionMode.ListIndex = dmFindFinalMolarity Then
        ' Check for Ctrl+C (copy)
        ' Other keystrokes are not allowed
        If KeyAscii <> 3 And KeyAscii <> 1 And KeyAscii <> 1 Then
            KeyAscii = 0
            Exit Sub
        End If
    End If

    TextBoxKeyPressHandler txtDilutionConcentrationFinal, KeyAscii, True, True, True, False, True, False, False, False, False, True

End Sub

Private Sub txtDilutionConcentrationFinal_Validate(Cancel As Boolean)
    RemoveNegativeValue txtDilutionConcentrationFinal
End Sub

Private Sub txtDilutionConcentrationInitial_Change()
    
    SynchronizeLinkedMolarities True
    
    If cboDilutionMode.ListIndex = dmFindInitialMolarity Then Exit Sub
    
    SynchronizeQuantitiesWithDLL mmtDilutionInitialConcentration
    
    FindDilutionValues
End Sub

Private Sub txtDilutionConcentrationInitial_GotFocus()
    HighlightOnFocus txtDilutionConcentrationInitial
End Sub

Private Sub txtDilutionConcentrationInitial_KeyPress(KeyAscii As Integer)
    If cboDilutionMode.ListIndex = dmFindInitialMolarity Then
        ' Check for Ctrl+C (copy)
        ' Other keystrokes are not allowed
        If KeyAscii <> 3 And KeyAscii <> 1 Then
            KeyAscii = 0
            Exit Sub
        End If
    End If
    
    TextBoxKeyPressHandler txtDilutionConcentrationInitial, KeyAscii, True, True, True, False, True, False, False, False, False, True

End Sub

Private Sub txtDilutionConcentrationInitial_Validate(Cancel As Boolean)
    RemoveNegativeValue txtDilutionConcentrationInitial
End Sub

Private Sub txtMassError_Change()
    ComputeMassRanges
End Sub

Private Sub txtMassError_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtMassError, KeyAscii, True, True, True, False, True, False, False, False, False, True, True
End Sub

Private Sub txtMassError_Validate(Cancel As Boolean)
     If IsNumeric(txtMassError) Then
        If CDblSafe(txtMassError) < 0 Then txtMassError = "0"
    End If
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

Private Sub txtVolume_Change()
    SynchronizeQuantitiesWithDLL mmtQuantityVolume
    mLastControlChanged = mmtQuantityVolume
End Sub

Private Sub txtVolume_GotFocus()
    HighlightOnFocus txtVolume
End Sub

Private Sub txtVolume_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtVolume, KeyAscii, True, True, True, False, True, False, False, False, False, True
End Sub

Private Sub txtDilutingSolventVolume_Change()
    HandleDilutingSolventVolumeChange
End Sub

Private Sub txtDilutingSolventVolume_GotFocus()
    HighlightOnFocus txtDilutingSolventVolume
End Sub

Private Sub txtDilutingSolventVolume_KeyPress(KeyAscii As Integer)
    If cboDilutionMode.ListIndex = dmFindRequiredDilutionVolumes Or _
       cboDilutionMode.ListIndex = dmFindRequiredTotalVolume Then
        ' Check for Ctrl+C (copy)
        ' Other keystrokes are not allowed
        If KeyAscii <> 3 And KeyAscii <> 1 Then
            KeyAscii = 0
            Exit Sub
        End If
    End If

    TextBoxKeyPressHandler txtDilutingSolventVolume, KeyAscii, True, True, True, False, True, False, False, False, False, True

End Sub

Private Sub txtDilutingSolventVolume_Validate(Cancel As Boolean)
    RemoveNegativeValue txtDilutingSolventVolume
End Sub

Private Sub txtStockSolutionVolume_Change()
    If cboDilutionMode.ListIndex = dmFindRequiredDilutionVolumes Then Exit Sub
    
    SynchronizeQuantitiesWithDLL mmtDilutionVolumeStock
    
    FindDilutionValues

End Sub

Private Sub txtStockSolutionVolume_GotFocus()
    HighlightOnFocus txtStockSolutionVolume
End Sub

Private Sub txtStockSolutionVolume_KeyPress(KeyAscii As Integer)
    If cboDilutionMode.ListIndex = dmFindRequiredDilutionVolumes Then
        ' Check for Ctrl+C (copy)
        If KeyAscii <> 3 And KeyAscii <> 1 Then
            KeyAscii = 0
            Exit Sub
        End If
    End If
    
    TextBoxKeyPressHandler txtStockSolutionVolume, KeyAscii, True, True, True, False, True, False, False, False, False, True

End Sub

Private Sub txtStockSolutionVolume_Validate(Cancel As Boolean)
    RemoveNegativeValue txtStockSolutionVolume
End Sub

Private Sub txtTotalVolume_Change()
    
    If cboDilutionMode.ListIndex = dmFindRequiredTotalVolume Then Exit Sub

    SynchronizeQuantitiesWithDLL mmtDilutionTotalVolume
    
    If mUpdatingTotalSolventVolume Then Exit Sub
    FindDilutionValues
    
End Sub

Private Sub txtTotalVolume_GotFocus()
    HighlightOnFocus txtTotalVolume
End Sub

Private Sub txtTotalVolume_KeyPress(KeyAscii As Integer)
    If cboDilutionMode.ListIndex = dmFindRequiredTotalVolume Then
        ' Check for Ctrl+C (copy)
        ' Other keystrokes are not allowed
        If KeyAscii <> 3 And KeyAscii <> 1 Then
            KeyAscii = 0
            Exit Sub
        End If
    End If

    TextBoxKeyPressHandler txtTotalVolume, KeyAscii, True, True, True, False, True, False, False, False, False, True

End Sub

Private Sub txtTotalVolume_Validate(Cancel As Boolean)
    RemoveNegativeValue txtTotalVolume
End Sub
