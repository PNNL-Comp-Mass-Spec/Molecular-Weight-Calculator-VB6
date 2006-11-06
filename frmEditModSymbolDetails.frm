VERSION 5.00
Begin VB.Form frmEditModSymbolDetails 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Amino Acid Modification Symbols Editor"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6210
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   6210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Tag             =   "15500"
   Begin VB.TextBox txtComment 
      Height          =   315
      Left            =   1560
      TabIndex        =   6
      Top             =   2160
      Width           =   4455
   End
   Begin VB.TextBox txtMass 
      Height          =   315
      Left            =   1560
      TabIndex        =   4
      Top             =   1680
      Width           =   1815
   End
   Begin VB.TextBox txtSymbol 
      Height          =   315
      Left            =   1560
      TabIndex        =   2
      Top             =   1200
      Width           =   1815
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4920
      TabIndex        =   7
      Tag             =   "4010"
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4920
      TabIndex        =   8
      Tag             =   "4020"
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove"
      Height          =   495
      Left            =   4920
      TabIndex        =   9
      Tag             =   "9020"
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label lblComment 
      Caption         =   "Commen&t"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Tag             =   "15140"
      Top             =   2190
      Width           =   1350
   End
   Begin VB.Label lblMass 
      Caption         =   "&Mass"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Tag             =   "15120"
      Top             =   1680
      Width           =   1350
   End
   Begin VB.Label lblSymbol 
      Caption         =   "&Symbol"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Tag             =   "15110"
      Top             =   1230
      Width           =   1350
   End
   Begin VB.Label lblInstructions 
      Caption         =   "Directions."
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Tag             =   "15250"
      Top             =   120
      Width           =   4575
   End
   Begin VB.Label lblHiddenButtonClickStatus 
      Caption         =   "-1"
      Height          =   255
      Left            =   5040
      TabIndex        =   10
      Top             =   1800
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "frmEditModSymbolDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    lblHiddenButtonClickStatus = BUTTON_CANCEL
    frmEditModSymbolDetails.Hide

End Sub

Private Sub cmdOK_Click()
    lblHiddenButtonClickStatus = BUTTON_OK
    frmEditModSymbolDetails.Hide
End Sub

Private Sub cmdRemove_Click()
    lblHiddenButtonClickStatus = BUTTON_RESET
    frmEditModSymbolDetails.Hide
End Sub

Private Sub Form_Activate()
    ' Put window in center of screen
    SizeAndCenterWindow Me, cWindowExactCenter, 6300, 3050
End Sub

Private Sub Form_Load()
    Me.Caption = LookupLanguageCaption(9000, "Amino Acid Modification Symbols Editor")
    
    CmdOK.Caption = LookupLanguageCaption(4010, "&Ok")
    cmdCancel.Caption = LookupLanguageCaption(4020, "&Cancel")
    cmdRemove.Caption = LookupLanguageCaption(9020, "&Remove")
    
    lblInstructions.Caption = LookupLanguageCaption(15750, "The amino acid symbol modification will be updated with the parameters below.  Select Remove to delete the modification symbol or Cancel to ignore any changes.")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    QueryUnloadFormHandler Me, Cancel, UnloadMode
End Sub

Private Sub txtComment_KeyPress(KeyAscii As Integer)
    ' Note: I'm purposely not using TextBoxKeyPressHandler for the Comment since I want to allow the user to type anything in the comment box
End Sub

Private Sub txtMass_Change()
    HighlightOnFocus txtMass
End Sub

Private Sub txtMass_GotFocus()
    HighlightOnFocus txtMass
End Sub

Private Sub txtMass_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtMass, KeyAscii, True, True, True, False, True
End Sub

Private Sub txtSymbol_GotFocus()
    HighlightOnFocus txtSymbol
End Sub

Private Sub txtSymbol_KeyPress(KeyAscii As Integer)
    ModSymbolKeyPressHandler txtSymbol, KeyAscii
End Sub
