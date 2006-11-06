VERSION 5.00
Begin VB.Form frmResults 
   Caption         =   "Form1"
   ClientHeight    =   6510
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   9285
   LinkTopic       =   "Form1"
   ScaleHeight     =   6510
   ScaleWidth      =   9285
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtResults 
      Height          =   5295
      Left            =   -240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   -120
      Width           =   7335
   End
End
Attribute VB_Name = "frmResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Function AddToResults(strNewText As String)
    txtResults = txtResults & vbCrLf & strNewText
    DoEvents
End Function

Private Sub Form_Resize()
    On Error Resume Next
    
    With txtResults
        .Top = 0
        .Left = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight
    End With
    
End Sub
