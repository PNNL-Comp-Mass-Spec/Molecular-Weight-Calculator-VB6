VERSION 5.00
Begin VB.Form frmSetValue 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Set Value"
   ClientHeight    =   1065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3945
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1065
   ScaleWidth      =   3945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Tag             =   "14500"
   Begin VB.TextBox txtEndVal 
      Height          =   285
      Left            =   1320
      TabIndex        =   3
      Text            =   "5"
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Tag             =   "4020"
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdSet 
      Caption         =   "&Set"
      Default         =   -1  'True
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Tag             =   "14510"
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtStartVal 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "5"
      Top             =   600
      Width           =   975
   End
   Begin VB.Label lblHiddenStatus 
      Caption         =   "Hidden Status"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblEndVal 
      Caption         =   "End Val"
      Height          =   435
      Left            =   1320
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblStartVal 
      Caption         =   "Start Val"
      Height          =   435
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1035
   End
End
Attribute VB_Name = "frmSetValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private blnCoerceDataWithinLimits As Boolean

Private mLowerLimit As Double
Private mUpperLimit As Double

Private localDefaultSeparationValue As Double

Public Sub SetLimits(blbLclCoerceDataWithinLimits As Boolean, Optional dblLowerLimit As Double = 0, Optional dblUpperLimit As Double = 100, Optional dblDefaultSeparationValue As Double = 1)
    blnCoerceDataWithinLimits = blbLclCoerceDataWithinLimits
    If blbLclCoerceDataWithinLimits Then
        mLowerLimit = dblLowerLimit
        mUpperLimit = dblUpperLimit
        localDefaultSeparationValue = dblDefaultSeparationValue
    End If
End Sub

Private Sub cmdCancel_Click()
    lblHiddenStatus = "Cancel"
    Me.Hide
End Sub

Private Sub cmdSet_Click()
    ' Must re-validate data since data is not validated if user presses Enter after changing a value
    If blnCoerceDataWithinLimits Then
        ValidateDualTextBoxes txtStartVal, txtEndVal, False, CDbl(mLowerLimit), CDbl(mUpperLimit), CDbl(localDefaultSeparationValue)
    End If
    
    lblHiddenStatus = "Ok"
    Me.Hide
End Sub

Private Sub Form_Load()
    SizeAndCenterWindow Me, cWindowUpperThird, , , False
    
    cmdSet.Caption = LookupLanguageCaption(14510, "&Set")
    cmdCancel.Caption = LookupLanguageCaption(4020, "&Cancel")
    Me.Caption = LookupLanguageCaption(14500, "Set Value")

End Sub

Private Sub txtEndVal_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtEndVal, KeyAscii, True, True, True, False, False, False, False, False, False, True, True
End Sub

Private Sub txtEndVal_Validate(Cancel As Boolean)
    If blnCoerceDataWithinLimits Then
        ValidateDualTextBoxes txtEndVal, txtEndVal, False, CDbl(mLowerLimit), CDbl(mUpperLimit), CDbl(localDefaultSeparationValue)
    End If
End Sub

Private Sub txtStartVal_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtStartVal, KeyAscii, True, True, True, False, False, False, False, False, False, True, True
End Sub

Private Sub txtStartVal_Validate(Cancel As Boolean)
    If blnCoerceDataWithinLimits Then
        ValidateDualTextBoxes txtStartVal, txtEndVal, False, CDbl(mLowerLimit), CDbl(mUpperLimit), CDbl(localDefaultSeparationValue)
    End If
End Sub
