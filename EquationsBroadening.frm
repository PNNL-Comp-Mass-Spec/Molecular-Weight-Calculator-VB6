VERSION 5.00
Begin VB.Form frmEquationsBroadening 
   Caption         =   "Extra-column broadening equations"
   ClientHeight    =   5790
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8760
   Icon            =   "EquationsBroadening.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   8760
   StartUpPosition =   3  'Windows Default
   Tag             =   "9500"
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   3360
      TabIndex        =   0
      Tag             =   "4010"
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   5925
      Left            =   0
      Picture         =   "EquationsBroadening.frx":08CA
      Top             =   0
      Width           =   8775
   End
End
Attribute VB_Name = "frmEquationsBroadening"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
    Me.Hide
End Sub

Private Sub Form_Activate()
    Me.Caption = LookupLanguageCaption(9500, "Extra-column broadening equations")
    cmdOK.Caption = LookupLanguageCaption(4010, "&Ok")
End Sub
