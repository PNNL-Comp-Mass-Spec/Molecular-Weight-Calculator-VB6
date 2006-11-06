VERSION 5.00
Begin VB.Form frmReportCompounds 
   Caption         =   "Report Compounds"
   ClientHeight    =   3660
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmReportCompounds.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3660
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCopy 
      Caption         =   "&Copy"
      Height          =   360
      HelpContextID   =   5000
      Left            =   600
      TabIndex        =   2
      Top             =   3120
      Width           =   1035
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   360
      HelpContextID   =   5000
      Left            =   2160
      TabIndex        =   1
      Top             =   3120
      Width           =   1035
   End
   Begin VB.ListBox lstCompounds 
      Height          =   2400
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   3135
   End
End
Attribute VB_Name = "frmReportCompounds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
    frmReportCompounds.Hide
End Sub

Private Sub cmdCopy_Click()
    Dim copytext$, x%
    
    Clipboard.Clear
    copytext$ = ""
    For x = 0 To lstCompounds.ListCount - 1
        ' MW not found, copy line to clipboard without any tabs
        copytext$ = copytext$ & Chr$(13) & Chr$(10) & lstCompounds.List(x)
        If x > 5000 Then Exit For
    Next x
    Clipboard.SetText copytext$
    
End Sub
