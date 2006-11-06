VERSION 5.00
Begin VB.Form frmAboutBox 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About MWT"
   ClientHeight    =   4320
   ClientLeft      =   315
   ClientTop       =   1695
   ClientWidth     =   5820
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "System"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   HelpContextID   =   6000
   Icon            =   "ABOUT.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4320
   ScaleWidth      =   5820
   Tag             =   "5500"
   Begin VB.PictureBox picIcon 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   4200
      Picture         =   "ABOUT.frx":08CA
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   4
      Top             =   1560
      Width           =   480
   End
   Begin VB.TextBox txtHTTP 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   10
      Text            =   "ABOUT.frx":0A07
      Top             =   2100
      Width           =   5055
   End
   Begin VB.TextBox txtEMail 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   9
      Text            =   "ABOUT.frx":0A9C
      Top             =   1560
      Width           =   3015
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      HelpContextID   =   6000
      Left            =   1800
      TabIndex        =   0
      Tag             =   "4010"
      Top             =   3855
      Width           =   1155
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
      TabIndex        =   11
      Top             =   840
      Width           =   2175
   End
   Begin VB.Line linLine1 
      BorderWidth     =   2
      X1              =   120
      X2              =   5160
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Label lblDate 
      Caption         =   "May 1, 1999"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   7
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label lblFreeware 
      Caption         =   "This program is Freeware; distribute freely"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Tag             =   "5510"
      Top             =   3060
      Width           =   5115
   End
   Begin VB.Label lblAuthor 
      Caption         =   "by Matthew Monroe"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Label lblMWT2 
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
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   600
      Width           =   3375
   End
   Begin VB.Label lblMWT 
      AutoSize        =   -1  'True
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
      Height          =   390
      Left            =   120
      TabIndex        =   1
      Top             =   150
      Width           =   5265
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
      Height          =   360
      Left            =   3840
      TabIndex        =   2
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label lblMicrosoft 
      Caption         =   "written using Microsoft Visual Basic v6.0 (SP5)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   3
      Top             =   3465
      Width           =   5205
   End
End
Attribute VB_Name = "frmAboutBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Purpose: Position controls on form
Private Sub PositionFormControls()
    Me.Caption = LookupLanguageCaption(5500, "About MWT")
    cmdOK.Caption = LookupLanguageCaption(4000, "Cl&ose")
    
    ' Position Objects
    lblMWT.Left = 120
    lblMWT.Top = 120
    LblMWT2.Left = 240
    LblMWT2.Top = 600
    lblVersion.Left = 3840
    lblVersion.Top = LblMWT2.Top
    lblBuild.Left = lblVersion.Left
    lblBuild.Top = LblMWT2.Top + 300
    lblAuthor.Left = lblMWT.Left
    lblAuthor.Top = 1200
    lblDate.Left = lblVersion.Left
    lblDate.Top = lblAuthor.Top
    With txtEMail
        .Left = lblMWT.Left
        .Top = 1560
        .Text = "Matt@AlchemistMatt.Com or AlchemistMatt@Yahoo.com"
    End With
    
    With txtHTTP
        .Left = lblMWT.Left
        .Top = 2100
        .Text = "http://www.alchemistmatt.com/" & vbCrLf & "http://ncrr.pnl.gov/software/" & vbCrLf & "http://come.to/alchemistmatt"
    End With
    
    lblFreeware.Left = lblMWT.Left
    lblFreeware.Top = 3050
    lblFreeware.Caption = LookupLanguageCaption(5510, "This program is Freeware; distribute freely")
    picIcon.Left = 4000
    picIcon.Top = 1560
    With linLine1
        .X1 = lblMWT.Left
        .X2 = 5200
        .Y1 = 3350
        .Y2 = 3350
    End With
    lblMicrosoft.Left = lblMWT.Left
    lblMicrosoft.Top = 3400
    cmdOK.Left = 1800
    cmdOK.Top = 3900
    
    lblVersion.Caption = "Version " & PROGRAM_VERSION
    lblBuild.Caption = "(Build " & App.Revision & ")"
    lblDate.Caption = PROGRAM_DATE

End Sub

Private Sub cmdOK_Click()
    Me.Hide
End Sub

Private Sub Form_Load()
    SizeAndCenterWindow Me, cWindowExactCenter, 5900, 4850, False
    PositionFormControls
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    QueryUnloadFormHandler Me, Cancel, UnloadMode
End Sub

