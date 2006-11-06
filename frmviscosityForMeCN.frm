VERSION 5.00
Begin VB.Form frmViscosityForMeCN 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Viscosity for Water/Acetonitrile Mixture"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   9360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Cl&ose"
      Height          =   360
      Left            =   7920
      TabIndex        =   13
      Tag             =   "4000"
      Top             =   1440
      Width           =   1155
   End
   Begin VB.CommandButton cmdDefaults 
      Caption         =   "Defaults"
      Height          =   360
      Left            =   5880
      TabIndex        =   12
      Tag             =   "16670"
      Top             =   1440
      Width           =   1800
   End
   Begin VB.PictureBox picChenHorvathEquation 
      Height          =   990
      Left            =   120
      Picture         =   "frmviscosityForMeCN.frx":0000
      ScaleHeight     =   930
      ScaleWidth      =   9075
      TabIndex        =   15
      Top             =   2160
      Width           =   9135
   End
   Begin VB.CommandButton cmdPlot 
      Caption         =   "Show Viscosity Plot"
      Height          =   360
      Left            =   5880
      TabIndex        =   11
      Tag             =   "16660"
      Top             =   960
      Width           =   1800
   End
   Begin VB.CommandButton cmdCopyViscosity 
      Caption         =   "Copy Viscosity"
      Default         =   -1  'True
      Height          =   360
      Left            =   5880
      TabIndex        =   10
      Tag             =   "16650"
      Top             =   480
      Width           =   1800
   End
   Begin VB.TextBox txtViscosity 
      Height          =   285
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "1"
      Top             =   1440
      Width           =   1095
   End
   Begin VB.ComboBox cboViscosityUnits 
      Height          =   315
      Left            =   3720
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Tag             =   "7040"
      Top             =   1440
      Width           =   1815
   End
   Begin VB.TextBox txtPercentAcetonitrile 
      Height          =   285
      Left            =   2520
      TabIndex        =   2
      Tag             =   "16600"
      Text            =   "20"
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox txtTemperature 
      Height          =   285
      Left            =   2520
      TabIndex        =   5
      Text            =   "25"
      Top             =   1080
      Width           =   1095
   End
   Begin VB.ComboBox cboTemperatureUnits 
      Height          =   315
      Left            =   3720
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Tag             =   "16510"
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label lblReference 
      Caption         =   "Chen, H; Horvath, CJ. J. Chromatography A, 1995, 705, 3"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   17
      Top             =   3480
      Width           =   6135
   End
   Begin VB.Label lblReference 
      Caption         =   "Thompson, JD; Carr, P. Analytical Chemistry, 2002, 74, 4150-4159."
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   16
      Top             =   3240
      Width           =   6135
   End
   Begin VB.Label lblChenHorvath 
      Caption         =   "Chen-Horvath Equation"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Tag             =   "16580"
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label lblDirections 
      Caption         =   "Enter the percent acetonitrile and temperature and the theoretical viscosity will be computed."
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Tag             =   "16550"
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label lblViscosity 
      Caption         =   "Solvent Viscosity"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Tag             =   "7230"
      Top             =   1470
      Width           =   2295
   End
   Begin VB.Label lblPercentAcetonitrileUnits 
      Caption         =   "%"
      Height          =   255
      Left            =   3240
      TabIndex        =   3
      Top             =   750
      Width           =   255
   End
   Begin VB.Label lblPercentOrganic 
      Caption         =   "Percent Acetonitrile"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Tag             =   "16560"
      Top             =   750
      Width           =   2295
   End
   Begin VB.Label lblTemperature 
      Caption         =   "Temperature"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Tag             =   "16570"
      Top             =   1110
      Width           =   2295
   End
End
Attribute VB_Name = "frmViscosityForMeCN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mUpdating As Boolean
Private mFormLoaded As Boolean

Private objViscosityPlot As CWSpectrumDLL.Spectrum

Private Sub ComputeViscosity()
    
    Dim eTemperatureUnits As utpUnitsTemperatureConstants
    Dim eViscosityUnits As uviUnitsViscosityConstants
    
    Dim dblTemperature As Double
    Dim dblPercentAcetonitrile As Double
    Dim dblViscosity As Double
    
    dblPercentAcetonitrile = CDblSafe(txtPercentAcetonitrile)
    dblTemperature = CDblSafe(txtTemperature)
    eTemperatureUnits = cboTemperatureUnits.ListIndex
    eViscosityUnits = cboViscosityUnits.ListIndex
    
    dblViscosity = Rnd(1) * 10
    dblViscosity = objMwtWin.CapFlow.ComputeMeCNViscosity(dblPercentAcetonitrile, dblTemperature, eTemperatureUnits, eViscosityUnits)
    
    txtViscosity = Round(dblViscosity, 6)

    UpdateViscosityPlot
End Sub

Private Sub CopyViscosityToCapillaryFlowForm()
    If IsNumeric(txtViscosity) Then
        frmCapillaryCalcs.cboCapValue(cccViscosityUnits).ListIndex = cboViscosityUnits.ListIndex
        frmCapillaryCalcs.txtCapValue(cctViscosity) = txtViscosity
    End If
End Sub

Private Sub PlotViscosityVsTemperature()
    If objViscosityPlot Is Nothing Then
        Set objViscosityPlot = New CWSpectrumDLL.Spectrum
        objViscosityPlot.SetCursorVisible False
    End If
    objViscosityPlot.ShowSpectrum
    UpdateViscosityPlot
End Sub

Public Sub SetDefaultValues()
    mUpdating = True
    txtTemperature = "25"
    cboTemperatureUnits.ListIndex = utpCelsius

    txtPercentAcetonitrile = "20"
    cboViscosityUnits.ListIndex = uviPoise
    
    mUpdating = False
    ComputeViscosity
    
End Sub

Public Sub SetViscosityValues(strViscosity As String, intViscosityUnits As Integer)

    txtViscosity = strViscosity
    If intViscosityUnits >= 0 And intViscosityUnits < cboViscosityUnits.ListCount Then
        cboViscosityUnits.ListIndex = intViscosityUnits
    End If
    
End Sub

Private Sub UpdateViscosityPlot()
    
    Const START_X As Long = 0
    Const END_X As Long = 100
    Const DELTA_X As Long = 1
    
    Dim lngIndex As Long
    
    Dim lngDataCount As Long
    Dim dblYValues() As Double
    
    Dim eTemperatureUnits As utpUnitsTemperatureConstants
    Dim dblTemperature As Double
    
    If Not objViscosityPlot Is Nothing Then
    
        lngDataCount = (END_X - START_X + 1) * DELTA_X
        ReDim dblYValues(lngDataCount - 1)
        
        dblTemperature = CDblSafe(txtTemperature)
        eTemperatureUnits = cboTemperatureUnits.ListIndex
        
        For lngIndex = START_X To END_X Step DELTA_X
            dblYValues(lngIndex) = objMwtWin.CapFlow.ComputeMeCNViscosity(CDbl(lngIndex), dblTemperature, eTemperatureUnits, cboViscosityUnits.ListIndex)
        Next lngIndex
        
        objViscosityPlot.SetSeriesPlotMode 1, pmLines, True
        objViscosityPlot.SetDataYOnly 1, dblYValues(), lngDataCount, START_X, DELTA_X, LookupLanguageCaption(16600, "Viscosity")
        objViscosityPlot.SetLabelXAxis LookupLanguageCaption(16610, "Percent Acetonitrile")
        objViscosityPlot.SetLabelYAxis LookupLanguageCaption(16600, "Viscosity") & " (" & cboViscosityUnits.Text & ")"
        
        objViscosityPlot.SetDisplayPrecisionX 0
        If cboViscosityUnits.ListIndex = uviPoise Then
            objViscosityPlot.SetDisplayPrecisionY 4
        Else
            objViscosityPlot.SetDisplayPrecisionY 2
        End If
        
    End If
End Sub

Private Sub PopulateComboBoxes()

    PopulateComboBox cboTemperatureUnits, True, "Celsius|Kelvin|Fahrenheit", 0   '7120
    PopulateComboBox cboViscosityUnits, True, "Poise [g/(cm-sec)]|centiPoise", 0   '7040

End Sub

Private Sub cboTemperatureUnits_Click()
    ComputeViscosity
End Sub

Private Sub cboViscosityUnits_Click()
    ComputeViscosity
End Sub

Private Sub cmdClose_Click()
    Me.Hide
End Sub

Private Sub cmdCopyViscosity_Click()
    CopyViscosityToCapillaryFlowForm
End Sub

Private Sub cmdDefaults_Click()
    SetDefaultValues
End Sub

Private Sub cmdPlot_Click()
    PlotViscosityVsTemperature
End Sub

Private Sub Form_Load()
    PopulateComboBoxes
    mFormLoaded = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    QueryUnloadFormHandler Me, Cancel, UnloadMode
End Sub

Private Sub txtPercentAcetonitrile_Change()
    ComputeViscosity
End Sub

Private Sub txtPercentAcetonitrile_GotFocus()
    HighlightOnFocus txtPercentAcetonitrile
End Sub

Private Sub txtPercentAcetonitrile_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtPercentAcetonitrile, KeyAscii, True, True
End Sub

Private Sub txtPercentAcetonitrile_LostFocus()
    ValidateTextboxValueDbl txtPercentAcetonitrile, 0, 100, 50
End Sub

Private Sub txtTemperature_Change()
    ComputeViscosity
End Sub

Private Sub txtTemperature_GotFocus()
    HighlightOnFocus txtTemperature
End Sub

Private Sub txtTemperature_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtTemperature, KeyAscii, True, True, True
End Sub

Private Sub txtTemperature_LostFocus()
    ValidateTextboxValueDbl txtTemperature, -500, 10000, 25
End Sub
