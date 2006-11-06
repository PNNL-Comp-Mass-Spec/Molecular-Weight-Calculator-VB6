VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmEditAbbrevDetails 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Editing Abbreviations"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6255
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Tag             =   "9000"
   Begin VB.TextBox txtComment 
      Height          =   405
      Left            =   1560
      TabIndex        =   10
      Top             =   3240
      Width           =   4455
   End
   Begin VB.TextBox txtOneLetterSymbol 
      Height          =   405
      Left            =   1560
      TabIndex        =   8
      Top             =   2760
      Width           =   1815
   End
   Begin VB.TextBox txtCharge 
      Height          =   405
      Left            =   1560
      TabIndex        =   6
      Top             =   2280
      Width           =   735
   End
   Begin VB.TextBox txtSymbol 
      Height          =   405
      Left            =   1560
      TabIndex        =   2
      Top             =   1200
      Width           =   1815
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4800
      TabIndex        =   11
      Tag             =   "4010"
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4800
      TabIndex        =   12
      Tag             =   "4020"
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove"
      Height          =   495
      Left            =   4800
      TabIndex        =   13
      Tag             =   "9020"
      Top             =   1080
      Width           =   1095
   End
   Begin RichTextLib.RichTextBox rtfFormula 
      Height          =   495
      Left            =   1560
      TabIndex        =   4
      Top             =   1680
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   873
      _Version        =   393217
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      TextRTF         =   $"frmEditAbbrevDetails.frx":0000
   End
   Begin VB.Label lblComment 
      Caption         =   "Commen&t"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Tag             =   "9195"
      Top             =   3270
      Width           =   1350
   End
   Begin VB.Label lblOneLetterSymbol 
      Caption         =   "&1 Letter"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Tag             =   "9190"
      Top             =   2790
      Width           =   1350
   End
   Begin VB.Label lblCharge 
      Caption         =   "Char&ge"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Tag             =   "9150"
      Top             =   2280
      Width           =   1350
   End
   Begin VB.Label lblFormula 
      Caption         =   "&Formula"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Tag             =   "9160"
      Top             =   1800
      Width           =   1350
   End
   Begin VB.Label lblSymbol 
      Caption         =   "&Symbol"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Tag             =   "9145"
      Top             =   1230
      Width           =   1350
   End
   Begin VB.Label lblInstructions 
      Caption         =   "Directions."
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label lblHiddenButtonClickStatus 
      Caption         =   "-1"
      Height          =   255
      Left            =   4920
      TabIndex        =   14
      Top             =   1920
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "frmEditAbbrevDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdCancel_Click()
    lblHiddenButtonClickStatus = BUTTON_CANCEL
    frmEditAbbrevDetails.Hide

End Sub

Private Sub cmdOK_Click()
    lblHiddenButtonClickStatus = BUTTON_OK
    frmEditAbbrevDetails.Hide
    
End Sub

Private Sub cmdRemove_Click()
    lblHiddenButtonClickStatus = BUTTON_RESET
    frmEditAbbrevDetails.Hide

End Sub

Private Sub Form_Activate()
    ' Put window in center of screen
    SizeAndCenterWindow Me, cWindowExactCenter, 6350, 4250
    
End Sub

Private Sub Form_Load()
    Me.Caption = LookupLanguageCaption(9000, "Editing Abbreviations")
    cmdOK.Caption = LookupLanguageCaption(4010, "&Ok")
    cmdCancel.Caption = LookupLanguageCaption(4020, "&Cancel")
    cmdRemove.Caption = LookupLanguageCaption(9020, "&Remove")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    QueryUnloadFormHandler Me, Cancel, UnloadMode
End Sub

Private Sub rtfFormula_Change()
    Dim saveloc As Integer
    
    saveloc = rtfFormula.SelStart
    rtfFormula.TextRTF = objMwtWin.TextToRTF(rtfFormula.Text)
    rtfFormula.SelStart = saveloc
    
End Sub

Private Sub rtfFormula_GotFocus()
    SetMostRecentTextBoxValue rtfFormula.Text
End Sub

Private Sub rtfFormula_KeyPress(KeyAscii As Integer)
    RTFBoxKeyPressHandler Me, rtfFormula, KeyAscii, False
End Sub

Private Sub txtCharge_Change()
    HighlightOnFocus txtCharge
End Sub

Private Sub txtCharge_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtCharge, KeyAscii, True, True, True, False, True
End Sub

Private Sub txtOneLetterSymbol_Change()
    HighlightOnFocus txtOneLetterSymbol
End Sub

Private Sub txtOneLetterSymbol_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtOneLetterSymbol, KeyAscii, False, False, False, True
End Sub

Private Sub txtOneLetterSymbol_Validate(Cancel As Boolean)
    If Len(txtOneLetterSymbol) > 1 Then
        txtOneLetterSymbol = Left(txtOneLetterSymbol, 1)
    End If
End Sub

Private Sub txtSymbol_GotFocus()
    HighlightOnFocus txtSymbol
End Sub

Private Sub txtSymbol_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtSymbol, KeyAscii, False, False, False, True
End Sub
