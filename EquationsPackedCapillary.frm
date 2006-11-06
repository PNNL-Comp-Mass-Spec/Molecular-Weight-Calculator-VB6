VERSION 5.00
Begin VB.Form frmEquationsPackedCapillary 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Equations for flow in a packed capillary"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8355
   Icon            =   "EquationsPackedCapillary.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   8355
   StartUpPosition =   3  'Windows Default
   Tag             =   "9600"
   Begin VB.PictureBox Picture3 
      Height          =   1455
      Left            =   4440
      Picture         =   "EquationsPackedCapillary.frx":08CA
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
      Picture         =   "EquationsPackedCapillary.frx":0FB0
      ScaleHeight     =   1395
      ScaleWidth      =   5475
      TabIndex        =   1
      Top             =   1680
      Width           =   5535
   End
   Begin VB.PictureBox Picture1 
      Height          =   1455
      Left            =   120
      Picture         =   "EquationsPackedCapillary.frx":1827
      ScaleHeight     =   1395
      ScaleWidth      =   4155
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
   Begin VB.Image Image1 
      Height          =   1245
      Left            =   120
      Picture         =   "EquationsPackedCapillary.frx":1E98
      Top             =   3240
      Width           =   5460
   End
End
Attribute VB_Name = "frmEquationsPackedCapillary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
    Me.Hide
End Sub

Private Sub Form_Activate()
    Me.Caption = LookupLanguageCaption(9600, "Equations for flow in a packed capillary")
    cmdOK.Caption = LookupLanguageCaption(4010, "&Ok")
End Sub
