VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmStrings 
   Caption         =   "String Reference Tables"
   ClientHeight    =   4905
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   6495
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdHide 
      Cancel          =   -1  'True
      Caption         =   "&Hide"
      Default         =   -1  'True
      Height          =   360
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Accepts changes made to abbreviations"
      Top             =   4440
      Width           =   1035
   End
   Begin MSFlexGridLib.MSFlexGrid grdLanguageStrings 
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   3625
      _Version        =   393216
      FixedRows       =   0
      FixedCols       =   0
      AllowUserResizing=   1
   End
   Begin MSFlexGridLib.MSFlexGrid grdMenuInfo 
      Height          =   2055
      Left            =   120
      TabIndex        =   1
      Top             =   2280
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   3625
      _Version        =   393216
      FixedRows       =   0
      FixedCols       =   0
      AllowUserResizing=   1
   End
   Begin MSFlexGridLib.MSFlexGrid grdLanguageStringsCrossRef 
      Height          =   2055
      Left            =   4440
      TabIndex        =   3
      Top             =   120
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   3625
      _Version        =   393216
      Rows            =   1
      FixedRows       =   0
      FixedCols       =   0
      AllowUserResizing=   1
   End
End
Attribute VB_Name = "frmStrings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdHide_Click()
    Me.Hide
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    QueryUnloadFormHandler Me, Cancel, UnloadMode
End Sub
