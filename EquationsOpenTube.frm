VERSION 5.00
Begin VB.Form frmEquationsOpenTube 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Equations for flow in an open tube"
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7725
   Icon            =   "EquationsOpenTube.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   7725
   StartUpPosition =   3  'Windows Default
   Tag             =   "9550"
   Begin VB.PictureBox Picture3 
      Height          =   1455
      Left            =   3480
      Picture         =   "EquationsOpenTube.frx":08CA
      ScaleHeight     =   1395
      ScaleWidth      =   3675
      TabIndex        =   3
      Top             =   120
      Width           =   3735
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   6480
      TabIndex        =   2
      Tag             =   "4010"
      Top             =   2160
      Width           =   1095
   End
   Begin VB.PictureBox Picture2 
      Height          =   1455
      Left            =   120
      Picture         =   "EquationsOpenTube.frx":0FA0
      ScaleHeight     =   1395
      ScaleWidth      =   5115
      TabIndex        =   1
      Top             =   1680
      Width           =   5175
   End
   Begin VB.PictureBox Picture1 
      Height          =   1455
      Left            =   120
      Picture         =   "EquationsOpenTube.frx":16FC
      ScaleHeight     =   1395
      ScaleWidth      =   3195
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
   Begin VB.Image Image1 
      Height          =   990
      Left            =   120
      Picture         =   "EquationsOpenTube.frx":1CA4
      Top             =   3240
      Width           =   5535
   End
End
Attribute VB_Name = "frmEquationsOpenTube"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
    Me.Hide
End Sub

Private Sub Form_Activate()
    Me.Caption = LookupLanguageCaption(9550, "Equations for flow in an open tube")
    cmdOK.Caption = LookupLanguageCaption(4010, "&Ok")
End Sub
