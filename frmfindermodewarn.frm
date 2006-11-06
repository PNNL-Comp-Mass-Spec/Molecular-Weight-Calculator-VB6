VERSION 5.00
Begin VB.Form frmFinderModeWarn 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Formula Finder"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6705
   HelpContextID   =   3045
   Icon            =   "frmFinderModeWarn.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   6705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Tag             =   "10000"
   Begin VB.OptionButton optWeightChoice 
      Caption         =   "Continue using &Average Weights."
      Height          =   255
      Index           =   2
      Left            =   480
      TabIndex        =   3
      Tag             =   "9770"
      Top             =   3600
      Width           =   5655
   End
   Begin VB.OptionButton optWeightChoice 
      Caption         =   "Always automatically switch to Isotopic &Weight mode."
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   2
      Tag             =   "9760"
      Top             =   3240
      Width           =   5655
   End
   Begin VB.OptionButton optWeightChoice 
      Caption         =   "Switch to &Isotopic Weight mode now."
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   1
      Tag             =   "9750"
      Top             =   2880
      Value           =   -1  'True
      Width           =   5655
   End
   Begin VB.CheckBox chkShowAgain 
      Caption         =   "&Stop showing this warning dialog."
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Tag             =   "9780"
      Top             =   4080
      Width           =   3855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Continue"
      Default         =   -1  'True
      Height          =   480
      Left            =   5040
      TabIndex        =   5
      Tag             =   "9720"
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label lblInstructions2 
      Caption         =   "Would you like to:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Tag             =   "9703"
      Top             =   2520
      Width           =   3375
   End
   Begin VB.Label lblInstructions 
      Caption         =   "Instructions"
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6495
   End
End
Attribute VB_Name = "frmFinderModeWarn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private intStopShowing As Integer

Public Sub DisplayInstructions(Optional blnlShowFormulaFinder As Boolean = True)
    Dim strMessage As String, strWindowCaption As String
    
    If blnlShowFormulaFinder Then
        ' Load correct text into strMessage
        strMessage = LookupLanguageCaption(9700, "The typical use of the Formula Finder feature is for when the monoisotopic mass (weight) of a compound is known (typically determined by Mass Spectrometry) and potential matching compounds are to be searched for.")
        strMessage = strMessage & vbCrLf & "     "
        strMessage = strMessage & LookupLanguageCaption(9701, "For example, a mass of 16.0312984 Daltons is measured for a compound containing Carbon and Hydrogen, and the possible empirical formula is desired.  Performing the search, with a weight tolerance of 5000 ppm results in three compounds, H2N, CH4, and O.  Within 500 ppm only CH4 matches, which is the correct match.")
        strMessage = strMessage & vbCrLf & "     "
        
        ' Set the proper window caption
        strWindowCaption = LookupLanguageCaption(10000, "Formula Finder")
    Else
        strMessage = LookupLanguageCaption(9705, "The typical use of the Fragmentation Modelling feature is for predicting the masses expected to be observed with a Mass Spectrometer when a peptide is ionized, enters the instrument, and fragments along the peptide backbone.")
        strMessage = strMessage & vbCrLf & "     "
        strMessage = strMessage & LookupLanguageCaption(9706, "The peptide typically fragments at each amide bond.  For example, the peptide Gly-Leu-Tyr will form the fragments Gly-Leu, Leu-Tyr, Gly, Leu, and Tyr.  Additionally, the cleavage of the amide bond can occur at differing locations, resulting in varying weights.")
        strMessage = strMessage & vbCrLf & "     "
        strWindowCaption = LookupLanguageCaption(12000, "Peptide Sequence Fragmentation Modelling")
    End If
    
    strMessage = strMessage & LookupLanguageCaption(9702, "To correctly use this feature, the program must be set to Isotopic Weight mode.  This can be done manually by choosing Edit Elements Table under the File menu, or the program can automatically switch to this mode for you.")
    
    lblInstructions.Caption = strMessage
    lblInstructions2.Caption = LookupLanguageCaption(9703, "Would you like to:")
    
    frmFinderModeWarn.Caption = strWindowCaption
End Sub

Private Sub cmdOK_Click()
    frmFinderModeWarn.Hide
    
End Sub

Private Sub Form_Activate()
    
    ' Put window in center of screen
    SizeAndCenterWindow Me, cWindowUpperThird, 7500
    
    intStopShowing = chkShowAgain.value
End Sub

Private Sub Form_GotFocus()
    optWeightChoice(0).SetFocus
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        optWeightChoice(0).value = True
        chkShowAgain.value = False
        cmdOK_Click
    End If
End Sub

Private Sub Form_Load()
    
    intStopShowing = chkShowAgain.value
    
    ' Set the first option as the default choice
    optWeightChoice(0).value = True
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    QueryUnloadFormHandler Me, Cancel, UnloadMode
End Sub

Private Sub optWeightChoice_Click(Index As Integer)
    If Index = 1 Then
        chkShowAgain.value = 1
    Else
        chkShowAgain.value = intStopShowing
    End If
        
End Sub
