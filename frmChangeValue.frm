VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmChangeValue 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Value"
   ClientHeight    =   1980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4545
   ControlBox      =   0   'False
   Icon            =   "frmChangeValue.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   4545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Tag             =   "8200"
   Begin RichTextLib.RichTextBox rtfValue 
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   873
      _Version        =   393217
      MultiLine       =   0   'False
      TextRTF         =   $"frmChangeValue.frx":08CA
   End
   Begin VB.TextBox txtValue 
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   1440
      Width           =   3015
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "&Reset to Default"
      Height          =   495
      Left            =   3360
      TabIndex        =   4
      Tag             =   "8210"
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Tag             =   "4020"
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Tag             =   "4010"
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblHiddenButtonClickStatus 
      Caption         =   "-1"
      Height          =   255
      Left            =   3480
      TabIndex        =   6
      Top             =   1680
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblInstructions 
      Caption         =   "Directions."
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "frmChangeValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    lblHiddenButtonClickStatus = BUTTON_CANCEL
    frmChangeValue.Hide

End Sub

Private Sub cmdOK_Click()
    lblHiddenButtonClickStatus = BUTTON_OK
    frmChangeValue.Hide
    
End Sub

Private Sub cmdReset_Click()
    lblHiddenButtonClickStatus = BUTTON_RESET
    frmChangeValue.Hide

End Sub

Private Sub Form_Activate()
    ' Put window in center of screen
    SizeAndCenterWindow Me, cWindowExactCenter, 4600, 2400
    
    If txtValue.Visible = True Then
        With txtValue
            .SetFocus
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
    Else
        With rtfValue
            .SetFocus
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
    End If

End Sub

Private Sub Form_Load()
    Me.Caption = LookupLanguageCaption(8200, "Change Value")
    CmdOK.Caption = LookupLanguageCaption(4010, "&Ok")
    cmdCancel.Caption = LookupLanguageCaption(4020, "&Cancel")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    QueryUnloadFormHandler Me, Cancel, UnloadMode
End Sub

Private Sub rtfValue_Change()
    Dim saveloc As Integer
    
    saveloc = rtfValue.SelStart
    rtfValue.TextRTF = objMwtWin.TextToRTF(rtfValue.Text)
    rtfValue.SelStart = saveloc
    
End Sub
