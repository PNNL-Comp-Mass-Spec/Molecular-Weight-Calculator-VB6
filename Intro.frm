VERSION 5.00
Begin VB.Form frmIntro 
   Caption         =   "Loading"
   ClientHeight    =   2475
   ClientLeft      =   2505
   ClientTop       =   930
   ClientWidth     =   5925
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   HelpContextID   =   500
   Icon            =   "INTRO.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2475
   ScaleWidth      =   5925
   Tag             =   "5700"
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   4560
      TabIndex        =   1
      Tag             =   "4030"
      Top             =   1200
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   4080
      TabIndex        =   0
      Tag             =   "4010"
      Top             =   1200
      Width           =   1035
   End
   Begin VB.Label lblBuild 
      BackStyle       =   0  'Transparent
      Caption         =   "(Build 35)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3840
      TabIndex        =   7
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label lblLoadStatus 
      AutoSize        =   -1  'True
      Height          =   315
      Left            =   120
      TabIndex        =   6
      Top             =   1695
      Width           =   3675
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "Version 5.07"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3840
      TabIndex        =   5
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label lblMWT 
      Caption         =   "Molecular Weight Calculator"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   5145
   End
   Begin VB.Label LblMWT2 
      Caption         =   "for Windows 9x/ME/NT/00/XP"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   3375
   End
   Begin VB.Label lblAuthor 
      Caption         =   "by Matthew Monroe"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   3615
   End
End
Attribute VB_Name = "frmIntro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const FORM_MIN_WIDTH = 6100
Private Const FORM_MIN_HEIGHT = 2850

Private Sub CheckLoadStatusLabelWidth()
    
    Dim lngLoadStatusWidth As Long
    
    ' Note that lblLoadStatus.AutoSize = True, and thus it automatically sets its width and height
    lngLoadStatusWidth = TextWidth(lblLoadStatus.Caption)
    
    If lngLoadStatusWidth < 3650 Then lngLoadStatusWidth = 3650
    If lngLoadStatusWidth > 15000 Then lngLoadStatusWidth = 15000
    
    lblLoadStatus.Width = lngLoadStatusWidth
        
    If lblLoadStatus.Width + 120 >= FORM_MIN_WIDTH Then Me.Width = lblLoadStatus.Width + 360
    If lblLoadStatus.Height > 675 Then Me.Height = lblLoadStatus.Top + lblLoadStatus.Height + 520
    
    If Me.Height + Me.Top > Screen.Height Then
        Me.Top = 0
    End If
        
End Sub

Private Sub cmdExit_Click()
    ' Unloading frmMain will result in all forms being unloaded and the program ending
    Unload frmMain
End Sub

Private Sub cmdOK_Click()
    gBlnLoadStatusOK = True
    frmMain.SetFocusToFormulaByIndex
    
    Unload Me
End Sub

Private Sub Form_Load()
    
    ' Put intro window in upper third of screen
    SizeAndCenterWindow Me, cWindowUpperThird, FORM_MIN_WIDTH, FORM_MIN_HEIGHT, False

    ' Position Objects
    lblMWT.Left = 120
    lblMWT.Top = 120
    LblMWT2.Left = 240
    LblMWT2.Top = 600
    lblVersion.Left = 3840
    lblVersion.Top = LblMWT2.Top
    lblBuild.Left = lblVersion.Left
    lblBuild.Top = lblVersion.Top + 300
    lblAuthor.Left = lblMWT.Left
    lblAuthor.Top = 1200
    lblLoadStatus.Left = lblMWT.Left
    lblLoadStatus.Top = 1600
    lblVersion.Caption = "Version " & PROGRAM_VERSION
    lblBuild.Caption = "(Build " & App.Revision & ")"
    
    ' Note that cmdOK and cmdExit are positioned on top of one another since only one is shown at a time
    With CmdOK
        .Left = lblVersion.Left
        .Top = lblAuthor.Top
        .Visible = False
        .Default = True
        .Cancel = True
    End With
    
    With cmdExit
        .Left = CmdOK.Left
        .Top = CmdOK.Top
        .Visible = False
        .Default = False
        .Cancel = False
    End With

End Sub

Private Sub lblLoadStatus_Change()
    CheckLoadStatusLabelWidth
End Sub
