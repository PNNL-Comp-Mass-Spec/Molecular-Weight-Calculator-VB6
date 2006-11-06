VERSION 5.00
Begin VB.Form frmMyMsgBox 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Caption"
   ClientHeight    =   1980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5850
   Icon            =   "MyMsgBox.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   5850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdButton 
      Caption         =   "Button4"
      Height          =   495
      HelpContextID   =   4002
      Index           =   4
      Left            =   4440
      TabIndex        =   4
      Tag             =   "0"
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "Button3"
      Height          =   495
      HelpContextID   =   4002
      Index           =   3
      Left            =   3000
      TabIndex        =   3
      Tag             =   "0"
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "Button2"
      Height          =   495
      HelpContextID   =   4002
      Index           =   2
      Left            =   1560
      TabIndex        =   2
      Tag             =   "0"
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "Button1"
      Default         =   -1  'True
      Height          =   495
      HelpContextID   =   4002
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Tag             =   "0"
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label lblMessage 
      Caption         =   "Message"
      Height          =   975
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "frmMyMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdButton_Click(Index As Integer)
    Me.Tag = Trim(Str(Index))
    Me.Hide
    
End Sub

Private Sub Form_Load()
    ' Put window in upper center of screen
    ' Do Not Use SizeAndCenterWindow since the MyMsgBox function repositions this form
    Me.Width = 4600
    Me.Height = 2400
    
End Sub
