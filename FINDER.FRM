VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmFinder 
   Caption         =   "Formula Finder"
   ClientHeight    =   6348
   ClientLeft      =   240
   ClientTop       =   1572
   ClientWidth     =   9828
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   7.8
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   HelpContextID   =   3050
   Icon            =   "FINDER.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6348
   ScaleWidth      =   9828
   Tag             =   "10000"
   Begin VB.CheckBox chkShowDeltaMass 
      Caption         =   "Show Delta Mass"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   77
      Tag             =   "10380"
      ToolTipText     =   "Note: Delta mass notation will change as the standard deviation mode is changed (F12 in the main window)"
      Top             =   5160
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "&Calculate"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   56
      Tag             =   "10310"
      ToolTipText     =   "Find the compounds that match the specified parameters"
      Top             =   5640
      Width           =   1035
   End
   Begin VB.CommandButton cmdAbort 
      Caption         =   "Abort"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   76
      Tag             =   "10455"
      Top             =   5640
      Width           =   1035
   End
   Begin VB.CommandButton cmdDisplayIsotopicAbundance 
      Caption         =   "&Display Iso Abundance"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   7800
      TabIndex        =   75
      Tag             =   "10345"
      Top             =   5040
      Width           =   1155
   End
   Begin VB.TextBox txtWeight 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   8280
      TabIndex        =   26
      Tag             =   "10270"
      Text            =   "# or Element or Abbrev."
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox txtWeight 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   8280
      TabIndex        =   21
      Tag             =   "10270"
      Text            =   "# or Element or Abbrev."
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton cmdFinderOptions 
      Caption         =   "Formul&a Finder Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   8040
      TabIndex        =   60
      Tag             =   "10300"
      ToolTipText     =   "Shortcut: Ctrl+O"
      Top             =   50
      Width           =   1395
   End
   Begin VB.CommandButton cmdCopyAsRTF 
      Caption         =   "Copy as RT&F"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   6480
      TabIndex        =   54
      Tag             =   "10330"
      Top             =   5040
      Width           =   1155
   End
   Begin RichTextLib.RichTextBox rtfResults 
      Height          =   2415
      Left            =   120
      TabIndex        =   64
      TabStop         =   0   'False
      Top             =   1560
      Width           =   3975
      _ExtentX        =   7006
      _ExtentY        =   4255
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      OLEDropMode     =   0
      TextRTF         =   $"FINDER.frx":08CA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtPercentMaxWeight 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3120
      TabIndex        =   47
      Tag             =   "10210"
      Text            =   "Enter Weight"
      Top             =   4200
      Width           =   1335
   End
   Begin VB.CheckBox chkPpmMode 
      Caption         =   "Ppm Mode"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   51
      Tag             =   "10370"
      Top             =   4920
      Width           =   1575
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Cop&y"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   6480
      TabIndex        =   58
      Tag             =   "10340"
      Top             =   5640
      Width           =   1035
   End
   Begin VB.TextBox txtHits 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8400
      TabIndex        =   4
      Tag             =   "10230"
      Text            =   "100"
      ToolTipText     =   "Maximum number of target compounds to find"
      Top             =   1200
      Width           =   615
   End
   Begin VB.TextBox txtWeightTolerance 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2760
      TabIndex        =   49
      Tag             =   "10220"
      Text            =   "0.5"
      ToolTipText     =   "Amount that target compound's weight can be from the target weight"
      Top             =   4800
      Width           =   1095
   End
   Begin VB.TextBox txtPercentTolerance 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2760
      TabIndex        =   50
      Tag             =   "10230"
      Text            =   "1"
      ToolTipText     =   "Amount that elemental percent compositions can be from the target percentage"
      Top             =   5040
      Width           =   735
   End
   Begin VB.TextBox txtMWT 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3120
      TabIndex        =   48
      Tag             =   "10210"
      Text            =   "Enter Weight"
      Top             =   4440
      Width           =   1335
   End
   Begin VB.TextBox txtMin 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   9
      Left            =   4440
      TabIndex        =   42
      Text            =   "0"
      Top             =   4440
      Width           =   495
   End
   Begin VB.TextBox txtMin 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   8
      Left            =   4440
      TabIndex        =   37
      Text            =   "0"
      Top             =   4080
      Width           =   495
   End
   Begin VB.TextBox txtMin 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   7
      Left            =   4440
      TabIndex        =   32
      Text            =   "0"
      Top             =   3720
      Width           =   495
   End
   Begin VB.TextBox txtMin 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   6
      Left            =   4440
      TabIndex        =   27
      Text            =   "0"
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox txtMin 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   4440
      TabIndex        =   22
      Text            =   "0"
      Top             =   3000
      Width           =   495
   End
   Begin VB.TextBox txtMin 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   4440
      TabIndex        =   17
      Text            =   "0"
      Top             =   2640
      Width           =   495
   End
   Begin VB.TextBox txtMin 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   4440
      TabIndex        =   13
      Text            =   "0"
      Top             =   2280
      Width           =   495
   End
   Begin VB.TextBox txtMin 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   4440
      TabIndex        =   9
      Text            =   "0"
      Top             =   1920
      Width           =   495
   End
   Begin VB.TextBox txtMin 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   4440
      TabIndex        =   5
      Text            =   "0"
      Top             =   1560
      Width           =   495
   End
   Begin VB.TextBox txtMin 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   4440
      TabIndex        =   0
      Text            =   "0"
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox txtMax 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   9
      Left            =   5040
      TabIndex        =   43
      Text            =   "10"
      Top             =   4440
      Width           =   495
   End
   Begin VB.TextBox txtMax 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   8
      Left            =   5040
      TabIndex        =   38
      Text            =   "10"
      Top             =   4080
      Width           =   495
   End
   Begin VB.TextBox txtMax 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   7
      Left            =   5040
      TabIndex        =   33
      Text            =   "10"
      Top             =   3720
      Width           =   495
   End
   Begin VB.TextBox txtMax 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   6
      Left            =   5040
      TabIndex        =   28
      Text            =   "10"
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox txtMax 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   5040
      TabIndex        =   23
      Text            =   "10"
      Top             =   3000
      Width           =   495
   End
   Begin VB.TextBox txtMax 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   5040
      TabIndex        =   18
      Text            =   "10"
      Top             =   2640
      Width           =   495
   End
   Begin VB.TextBox txtMax 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   5040
      TabIndex        =   14
      Text            =   "10"
      Top             =   2280
      Width           =   495
   End
   Begin VB.TextBox txtMax 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   5040
      TabIndex        =   10
      Text            =   "10"
      Top             =   1920
      Width           =   495
   End
   Begin VB.TextBox txtMax 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   5040
      TabIndex        =   6
      Text            =   "10"
      Top             =   1560
      Width           =   495
   End
   Begin VB.TextBox txtMax 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   5040
      TabIndex        =   1
      Text            =   "10"
      Top             =   1200
      Width           =   495
   End
   Begin VB.CheckBox chkElements 
      Caption         =   "Element 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   5640
      TabIndex        =   2
      Top             =   1200
      Width           =   1700
   End
   Begin VB.CheckBox chkElements 
      Caption         =   "Element 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   5640
      TabIndex        =   7
      Top             =   1560
      Width           =   1700
   End
   Begin VB.CheckBox chkElements 
      Caption         =   "Element 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   5640
      TabIndex        =   11
      Top             =   1920
      Width           =   1700
   End
   Begin VB.CheckBox chkElements 
      Caption         =   "Element 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   5640
      TabIndex        =   15
      Top             =   2280
      Width           =   1700
   End
   Begin VB.CheckBox chkElements 
      Caption         =   "Element 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   5640
      TabIndex        =   19
      Top             =   2640
      Width           =   1700
   End
   Begin VB.CheckBox chkElements 
      Caption         =   "Element 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   5640
      TabIndex        =   24
      Top             =   3000
      Width           =   1700
   End
   Begin VB.CheckBox chkElements 
      Caption         =   "Element 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   5640
      TabIndex        =   29
      Top             =   3360
      Width           =   1700
   End
   Begin VB.CheckBox chkElements 
      Caption         =   "Element 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   5640
      TabIndex        =   34
      Top             =   3720
      Width           =   1700
   End
   Begin VB.CheckBox chkElements 
      Caption         =   "Element 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   5640
      TabIndex        =   39
      Top             =   4080
      Width           =   1700
   End
   Begin VB.CheckBox chkElements 
      Caption         =   "Element 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   5640
      TabIndex        =   44
      Top             =   4440
      Width           =   1700
   End
   Begin VB.TextBox txtPercent 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   7440
      TabIndex        =   3
      Tag             =   "10260"
      Text            =   "Percent"
      Top             =   1200
      Width           =   780
   End
   Begin VB.TextBox txtPercent 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   7440
      TabIndex        =   8
      Tag             =   "10260"
      Text            =   "Percent"
      Top             =   1560
      Width           =   780
   End
   Begin VB.TextBox txtPercent 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   7440
      TabIndex        =   12
      Tag             =   "10260"
      Text            =   "Percent"
      Top             =   1920
      Width           =   780
   End
   Begin VB.TextBox txtPercent 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   7440
      TabIndex        =   16
      Tag             =   "10260"
      Text            =   "Percent"
      Top             =   2280
      Width           =   780
   End
   Begin VB.TextBox txtPercent 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   7440
      TabIndex        =   20
      Tag             =   "10260"
      Text            =   "Percent"
      Top             =   2640
      Width           =   780
   End
   Begin VB.TextBox txtPercent 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   7440
      TabIndex        =   25
      Tag             =   "10260"
      Text            =   "Percent"
      Top             =   3000
      Width           =   780
   End
   Begin VB.TextBox txtPercent 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   6
      Left            =   7440
      TabIndex        =   30
      Tag             =   "10260"
      Text            =   "Percent"
      Top             =   3360
      Width           =   780
   End
   Begin VB.TextBox txtPercent 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   7
      Left            =   7440
      TabIndex        =   35
      Tag             =   "10260"
      Text            =   "Percent"
      Top             =   3720
      Width           =   780
   End
   Begin VB.TextBox txtPercent 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   8
      Left            =   7440
      TabIndex        =   40
      Tag             =   "10260"
      Text            =   "Percent"
      Top             =   4080
      Width           =   780
   End
   Begin VB.TextBox txtPercent 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   9
      Left            =   7440
      TabIndex        =   45
      Tag             =   "10260"
      Text            =   "Percent"
      Top             =   4440
      Width           =   780
   End
   Begin VB.TextBox txtWeight 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   6
      Left            =   8280
      TabIndex        =   31
      Tag             =   "10270"
      Text            =   "# or Element or Abbrev."
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox txtWeight 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   7
      Left            =   8280
      TabIndex        =   36
      Tag             =   "10270"
      Text            =   "# or Element or Abbrev."
      Top             =   3720
      Width           =   975
   End
   Begin VB.TextBox txtWeight 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   8
      Left            =   8280
      TabIndex        =   41
      Tag             =   "10270"
      Text            =   "# or Element or Abbrev."
      Top             =   4080
      Width           =   975
   End
   Begin VB.TextBox txtWeight 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   9
      Left            =   8280
      TabIndex        =   46
      Tag             =   "10270"
      Text            =   "# or Element or Abbrev."
      Top             =   4440
      Width           =   975
   End
   Begin VB.OptionButton optType 
      Caption         =   "Match &Molecular Weight"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   52
      Tag             =   "10350"
      Top             =   5400
      Value           =   -1  'True
      Width           =   2955
   End
   Begin VB.OptionButton optType 
      Caption         =   "Match &Percent Compositions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   55
      Tag             =   "10360"
      Top             =   5760
      Width           =   3000
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Prin&t ..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   5160
      TabIndex        =   57
      Tag             =   "10320"
      Top             =   5640
      Width           =   1035
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Cl&ose"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   7800
      TabIndex        =   59
      Tag             =   "4000"
      Top             =   5640
      Width           =   1035
   End
   Begin VB.ListBox lstResults 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2736
      Left            =   120
      TabIndex        =   62
      Top             =   1200
      Width           =   3855
   End
   Begin VB.Label lblMin 
      Caption         =   "Min"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   63
      Tag             =   "10060"
      Top             =   720
      Width           =   495
   End
   Begin VB.Label lblWeight 
      Caption         =   "Atomic Wt."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8400
      TabIndex        =   70
      Tag             =   "10105"
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label lblPercentMaxWeight 
      Caption         =   "Maximum Weight of Formula:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   74
      Tag             =   "10020"
      Top             =   4200
      Width           =   2895
   End
   Begin VB.Label lblWeightMode 
      Caption         =   "(using average atomic weights)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   73
      Top             =   720
      Width           =   3735
   End
   Begin VB.Label lblPercentCompleted 
      Alignment       =   2  'Center
      Caption         =   "Percent Completed"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8280
      TabIndex        =   72
      Tag             =   "10110"
      Top             =   1560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblHits 
      Caption         =   "Max Hits"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8400
      TabIndex        =   71
      Tag             =   "10100"
      Top             =   720
      Width           =   975
   End
   Begin VB.Label lblPercent 
      Alignment       =   2  'Center
      Caption         =   "Percent"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7320
      TabIndex        =   69
      Tag             =   "10090"
      Top             =   720
      Width           =   975
   End
   Begin VB.Label lblElement 
      Caption         =   "Element"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   68
      Tag             =   "10080"
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label lblWtTolerance 
      Caption         =   "Weight Tolerance:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   66
      Tag             =   "10040"
      Top             =   4800
      Width           =   2535
   End
   Begin VB.Label lblPercentTolerance 
      Caption         =   "Percent Tolerance:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   67
      Tag             =   "10050"
      Top             =   5040
      Width           =   2535
   End
   Begin VB.Label lblInstruct 
      Caption         =   $"FINDER.frx":094E
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   53
      Tag             =   "10010"
      Top             =   0
      Width           =   6735
   End
   Begin VB.Label lblMWT 
      Caption         =   "Molecular Weight of Target:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   61
      Tag             =   "10030"
      Top             =   4440
      Width           =   2895
   End
   Begin VB.Label lblMax 
      Caption         =   "Max"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      TabIndex        =   65
      Tag             =   "10070"
      Top             =   720
      Width           =   615
   End
   Begin VB.Menu mnuRightClick 
      Caption         =   "RightClickMenu"
      Begin VB.Menu mnuRightClickUndo 
         Caption         =   "&Undo"
      End
      Begin VB.Menu mnuRightClickSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRightClickCut 
         Caption         =   "Cu&t"
      End
      Begin VB.Menu mnuRightClickCopy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnuRightClickPaste 
         Caption         =   "&Paste"
      End
      Begin VB.Menu mnuRightClickDelete 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnuRightClickSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRightClickSelectAll 
         Caption         =   "Select &All"
      End
   End
End
Attribute VB_Name = "frmFinder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'  The x,0 element of codearray stores a pointer to the original line the code comes from
'  Then, x,1 through x,MAX_MATCHINGELEMENTS*2 is the code
'  pointerarray stores a pointer to a specific line in codearray, so that swapping while sorting is extremely fast

Private Const MAX_MATCHINGELEMENTS = 10
Private Const MAX_HITS = 15000

Private lngRecursiveFunctionCallCount As Long
Private mDecimalSeparator As String

Private Type udtElementNumType
    H As Integer
    C As Integer
    Si As Integer
    N As Integer
    P As Integer
    O As Integer
    S As Integer
    Cl As Integer
    I As Integer
    F As Integer
    Br As Integer
    Other As Integer
End Type

Private Type udtResultStatsType
    Charge As Integer
    Mass As Double
    MZ As Double
    DeltaMass As Double
End Type

Private Function AbsMax(ByRef intValueMin As Integer, ByRef intValueMax As Integer) As Integer
    ' Takes the absolute value of both values (updating the result)
    ' Makes sure the minimum value is in intValueMin and the maximum in intValueMax
    
    Dim intSwapVal As Integer
    
    intValueMin = Abs(intValueMin)
    intValueMax = Abs(intValueMax)
    
    If intValueMin > intValueMax Then
        intSwapVal = intValueMax
        intValueMax = intValueMin
        intValueMin = intSwapVal
    End If
    
    AbsMax = intValueMax

End Function

Private Sub CheckMtoZWithTarget(dblTotalMass As Double, intTotalCharge As Single, dblTargetMass As Double, dblMassTolerance As Double, intMultipleMtoZCharge As Integer, blnChargeOK As Boolean)
    ' Warning: This sub assumes that intMultipleMtoZCharge > 0
    Dim dblMtoZ As Double, dblOriginalMtoZ As Double
        
    ' The original target is the target m/z; assure this compound has that m/z
    If intTotalCharge <> 0 Then
        dblMtoZ = Abs(dblTotalMass / intTotalCharge)
    Else
        dblMtoZ = 0
    End If
        
    ' blnChargeOK is True going into this sub
    dblOriginalMtoZ = dblTargetMass / intMultipleMtoZCharge
    If dblMtoZ < dblOriginalMtoZ - dblMassTolerance Or dblMtoZ > dblOriginalMtoZ + dblMassTolerance Then
        ' dblMtoZ is not wither tolerance of dblOriginalMtoZ, so don't print
        blnChargeOK = False
    End If

End Sub

Private Sub CompoundToCode(ByVal dblTargetWeight As Double, ByRef strCodeString() As String, ByRef udtResultsStats() As udtResultStatsType, ByRef intPointerArray() As Integer)

    ' Precedence convert table for strCodeString:
    '  C1_ C2_ C3_ C4_   a   z    1,  2,  3...
    '   1   2   3   4    5  30    31  32  33
    ' I convert the above numbers to characters using Chr() and add a Chr(0) in front of each one so that ...
    ' if a number is larger than 226, then I use two characters:
    ' 226 = Chr(0) + Chr(226+29) = Chr(0) + Chr(255)
    ' 227 = Chr(CIntSafeDbl(226+29/255)) + Chr(226 + 29 Mod 255)

    Dim strCurrentLetter As String, strNumString As String, strCompound As String, strRemainder As String
    Dim intXTrack As Integer, intCharLoc As Integer, intNumberLength As Integer
    Dim intLineLength As Integer, intAddon As Integer
    Dim intCurrentCharIndex As Integer, intCharAscii As Integer, intCurrentLine As Integer
    Dim lngCurrentNum As Long, intNumSize As Integer
            
On Error GoTo CompoundToCodeErrorHandler

    'Erase strCodeString
    For intCurrentLine = 0 To lstResults.ListCount - 1
        intPointerArray(intCurrentLine) = intCurrentLine
        
        strCompound = lstResults.List(intCurrentLine)
        intXTrack = -1
        intLineLength = Len(strCompound)
        For intCurrentCharIndex = 1 To intLineLength
            ' Check to see if we're at charge (^) or the MW= section yet
            strCurrentLetter = UCase(Mid(strCompound, intCurrentCharIndex, 1))
            If strCurrentLetter = "^" Or _
               strCurrentLetter = vbTab Or _
               strCurrentLetter = " " Then
                ' Now fill udtResultStats()
                If strCurrentLetter = "^" Then
                    ' find entire number after caret
                    intNumberLength = 1
                    Do While IsNumeric(Mid(strCompound, intCurrentCharIndex + intNumberLength, 1)) Or _
                             Mid(strCompound, intCurrentCharIndex + intNumberLength, 1) = "-"
                        intNumberLength = intNumberLength + 1
                    Loop
                    'dblResultStats(intCurrentLine, 0) = Val(Mid(strCompound, intCurrentCharIndex + 1, intNumberLength - 1))
                    udtResultsStats(intCurrentLine).Charge = Val(Mid(strCompound, intCurrentCharIndex + 1, intNumberLength - 1))
                    intCurrentCharIndex = intCurrentCharIndex + intNumberLength
                End If
                strRemainder = Mid(strCompound, intCurrentCharIndex)
                
                intCharLoc = InStr(strRemainder, gMWAbbreviation & "=")
                If intCharLoc > 0 Then
                    ' Now find MW (or foreign equivalent in gMWAbbreviation)
                    intCharLoc = intCharLoc + Len(gMWAbbreviation & "=")
                    'dblResultStats(intCurrentLine, 1) = ParseNum(Mid(strRemainder, intCharLoc), intNumSize)
                    udtResultsStats(intCurrentLine).Mass = ParseNum(Mid(strRemainder, intCharLoc), intNumSize)
                    udtResultsStats(intCurrentLine).DeltaMass = Abs(udtResultsStats(intCurrentLine).Mass - dblTargetWeight)
                    strRemainder = Mid(strRemainder, intCharLoc + intNumSize)
                Else
                    'dblResultStats(intCurrentLine, 1) = 0
                    udtResultsStats(intCurrentLine).Mass = 0
                    udtResultsStats(intCurrentLine).DeltaMass = 0
                    intNumSize = 0
                End If
                
                intCharLoc = InStr(strRemainder, "m/z")
                If intCharLoc > 0 Then
                    ' Now find m/z
                    intCharLoc = intCharLoc + 3
                    'dblResultStats(intCurrentLine, 2) = ParseNum(Mid(strRemainder, intCharLoc), intNumSize)
                    udtResultsStats(intCurrentLine).MZ = ParseNum(LTrim(Mid(strRemainder, intCharLoc)), intNumSize)
                Else
                    'dblResultStats(intCurrentLine, 2) = 0
                    udtResultsStats(intCurrentLine).MZ = 0
                End If
                
                Exit For
            End If
            intAddon = 1
            intCharAscii = Asc(strCurrentLetter)
            If (intCharAscii >= 65 And intCharAscii <= 90) Then
                ' at a letter
                intXTrack = intXTrack + 1
                If Mid(strCompound, intCurrentCharIndex + 2, 1) = "_" Then
                    ' At a custom element, give it a value of 1 to 4
                    ' Also, need to bump up intCurrentCharIndex by 2
                    intAddon = 3
                    ' At a custom element, give it a value of Chr(0) to Chr(3), thus subtract 1 from x+1 place in string
                    strCodeString(intCurrentLine) = strCodeString(intCurrentLine) + Chr(0) + _
                                           Chr(Val(Mid(strCompound, intCurrentCharIndex + 1, 1)))
                Else
                    ' 65 is the ascii code for the letter a
                    ' Thus, 65-60 = 5; subtract 60 since custom elements are Chr(0) to Chr(3), so A is Chr(5)
                    strCodeString(intCurrentLine) = strCodeString(intCurrentLine) + Chr(0) + _
                                           Chr(intCharAscii - 60)
                End If
            Else
                ' at a number, since only letters and numbers in Formula Finder results!
                lngCurrentNum = 0
                strNumString = strCurrentLetter
                Do While IsNumeric(Mid(strCompound, intCurrentCharIndex + intAddon, 1)) = True
                    strNumString = strNumString & Mid(strCompound, intCurrentCharIndex + 1, 1)
                    intAddon = intAddon + 1
                Loop
                lngCurrentNum = Val(strNumString)
                If lngCurrentNum > 0 Then
                    intXTrack = intXTrack + 1
                    ' Add 30 so that the number 1 is represented as 31
                    If lngCurrentNum <= 225 Then
                        strCodeString(intCurrentLine) = strCodeString(intCurrentLine) + Chr(0) + _
                                               Chr(lngCurrentNum + 30)
                    Else
                        ' lngCurrentNum must be less than 65025 !
                        strCodeString(intCurrentLine) = strCodeString(intCurrentLine) + _
                                               Chr(CIntSafeDbl((lngCurrentNum + 30) / 255)) + _
                                               Chr((lngCurrentNum + 30) Mod 255 + 1)
                    End If
                End If
            End If
            If intXTrack = MAX_MATCHINGELEMENTS * 2 And intCurrentCharIndex < intLineLength Then
                   ' Too many things to compare (over MAX_MATCHINGELEMENTS elements) -- Shouldn't occur since users are limited to MAX_MATCHINGELEMENTS found elements/abbreviations
                   ' Save the rest of the string in strCodeString
                   strCodeString(intCurrentLine) = strCodeString(intCurrentLine) & "&" & _
                                          Mid(strCompound, intCurrentCharIndex + 1)
            End If
            intCurrentCharIndex = intCurrentCharIndex + intAddon - 1
        Next intCurrentCharIndex
        
    Next intCurrentLine
        
    Exit Sub

CompoundToCodeErrorHandler:
    GeneralErrorHandler "frmFinder|CompoundToCode", Err.Number, Err.Description

End Sub

Private Function ConstructAndVerifyCompound(ByRef strConstruct As String, count1 As Long, count2 As Long, count3 As Long, count4 As Long, count5 As Long, count6 As Long, count7 As Long, count8 As Long, count9 As Long, count10 As Long, elem1$, elem2$, elem3$, elem4$, elem5$, elem6$, elem7$, elem8$, elem9$, elem10$, dblTotalWeight As Double, dblTargetWeight As Double, dblWeightTolerance As Double, sngTotalCharge As Single, intMultipleMtoZCharge As Integer) As Boolean
    ' Returns 0 if compound has too many hydrogens AND hydrogen checking is on,
    '  otherwise returns 1
    '
    ' common function to both molecular weight and percent composition matching
    '
    ' strConstruct is passed by reference and is modified in this subroutine
    
    Dim X As Integer, Y As Integer
    Dim swapVal As Integer, blnHOK As Boolean, blnChargeOK As Boolean
    Dim strSwap As String
    Dim strFormulaElem(10) As String
    Dim intFormulaCount(10) As Integer
    Dim EmpResultSymbols(9) As String   ' Store the elements and abbreviations of the found formula so that they can be properly sequenced according to empirical formula conventions
    Dim EmpResultCount(9) As Integer    ' Store the number of each element and abbreviation for empirical formula conversion
    
    ' Counters for elements of interest (hydrogen, carbon, silicon, nitrogen, phosphorus, chlorine, iodine, flourine, bromine, and other)
    Dim udtElementNum As udtElementNumType
    
On Error GoTo ConstructAndVerifyCompoundErrorHandler

    If Not cChkBox(frmFinderOptions.chkSort) Then
        ' Don't convert to empirical formula or sort
        If count1 <> 0 Then strConstruct = strConstruct & elem1$
        If count1 > 1 Then strConstruct = strConstruct & LTrim(Str(count1))
        If count2 <> 0 Then strConstruct = strConstruct & elem2$
        If count2 > 1 Then strConstruct = strConstruct & LTrim(Str(count2))
        If count3 <> 0 Then strConstruct = strConstruct & elem3$
        If count3 > 1 Then strConstruct = strConstruct & LTrim(Str(count3))
        If count4 <> 0 Then strConstruct = strConstruct & elem4$
        If count4 > 1 Then strConstruct = strConstruct & LTrim(Str(count4))
        If count5 <> 0 Then strConstruct = strConstruct & elem5$
        If count5 > 1 Then strConstruct = strConstruct & LTrim(Str(count5))
        If count6 <> 0 Then strConstruct = strConstruct & elem6$
        If count6 > 1 Then strConstruct = strConstruct & LTrim(Str(count6))
        If count7 <> 0 Then strConstruct = strConstruct & elem7$
        If count7 > 1 Then strConstruct = strConstruct & LTrim(Str(count7))
        If count8 <> 0 Then strConstruct = strConstruct & elem8$
        If count8 > 1 Then strConstruct = strConstruct & LTrim(Str(count8))
        If count9 <> 0 Then strConstruct = strConstruct & elem9$
        If count9 > 1 Then strConstruct = strConstruct & LTrim(Str(count9))
        If count10 <> 0 Then strConstruct = strConstruct & elem10$
        If count10 > 1 Then strConstruct = strConstruct & LTrim(Str(count10))
    Else
        For X = 0 To 9
            EmpResultSymbols(X) = ""
            EmpResultCount(X) = 0
        Next X
        ' Convert to empirical formula and sort
        If count1 <> 0 Then EmpResultSymbols(0) = elem1$
        If count1 > 1 Then EmpResultCount(0) = count1
        If count2 <> 0 Then EmpResultSymbols(1) = elem2$
        If count2 > 1 Then EmpResultCount(1) = count2
        If count3 <> 0 Then EmpResultSymbols(2) = elem3$
        If count3 > 1 Then EmpResultCount(2) = count3
        If count4 <> 0 Then EmpResultSymbols(3) = elem4$
        If count4 > 1 Then EmpResultCount(3) = count4
        If count5 <> 0 Then EmpResultSymbols(4) = elem5$
        If count5 > 1 Then EmpResultCount(4) = count5
        If count6 <> 0 Then EmpResultSymbols(5) = elem6$
        If count6 > 1 Then EmpResultCount(5) = count6
        If count7 <> 0 Then EmpResultSymbols(6) = elem7$
        If count7 > 1 Then EmpResultCount(6) = count7
        If count8 <> 0 Then EmpResultSymbols(7) = elem8$
        If count8 > 1 Then EmpResultCount(7) = count8
        If count9 <> 0 Then EmpResultSymbols(8) = elem9$
        If count9 > 1 Then EmpResultCount(8) = count9
        If count10 <> 0 Then EmpResultSymbols(9) = elem10$
        If count10 > 1 Then EmpResultCount(9) = count10
        ' First find C
        For X = 0 To 9
            If EmpResultSymbols(X) = "C" Then
                strConstruct = strConstruct & EmpResultSymbols(X)       ' i.e. C
                If EmpResultCount(X) > 1 Then strConstruct = strConstruct & _
                                                         LTrim(Str(EmpResultCount(X)))
                ' blank this EmpResultSymbols so it doesn't get repeated
                EmpResultSymbols(X) = ""
                Exit For
            End If
        Next X
        For X = 0 To 9
            If EmpResultSymbols(X) = "H" Then
                strConstruct = strConstruct & EmpResultSymbols(X)       ' i.e. H
                If EmpResultCount(X) > 1 Then strConstruct = strConstruct & _
                                                         LTrim(Str(EmpResultCount(X)))
                ' blank this EmpResultSymbols so it doesn't get repeated
                EmpResultSymbols(X) = ""
                Exit For
            End If
        Next X
        
        ' Alphabatize the remaining elements/abbreviations via a simple bubble sort
        For Y = 9 To 1 Step -1       ' Sort from end to start
            For X = 0 To Y - 1
                If EmpResultSymbols(X) > EmpResultSymbols(X + 1) Then
                    ' Swap the elements/abbreviations
                    strSwap = EmpResultSymbols(X)
                    EmpResultSymbols(X) = EmpResultSymbols(X + 1)
                    EmpResultSymbols(X + 1) = strSwap
                    ' and their EmpResultCount values
                    swapVal = EmpResultCount(X)
                    EmpResultCount(X) = EmpResultCount(X + 1)
                    EmpResultCount(X + 1) = swapVal
                End If
            Next X
        Next Y
        
        ' Now insert the alphabatized elements into the string
        For X = 0 To 9
            If EmpResultSymbols(X) <> "" Then
                strConstruct = strConstruct & EmpResultSymbols(X)
                If EmpResultCount(X) > 1 Then strConstruct = strConstruct & _
                                                         LTrim(Str(EmpResultCount(X)))
            End If
        Next X

    End If
    
    
    ' Verify hydrogens if requested
    If cChkBox(frmFinderOptions.chkVerifyHydrogens) Or cChkBox(frmFinderOptions.chkFindTargetMtoZ) Then
        intFormulaCount(1) = count1
        strFormulaElem(1) = elem1$
        intFormulaCount(2) = count2
        strFormulaElem(2) = elem2$
        intFormulaCount(3) = count3
        strFormulaElem(3) = elem3$
        intFormulaCount(4) = count4
        strFormulaElem(4) = elem4$
        intFormulaCount(5) = count5
        strFormulaElem(5) = elem5$
        intFormulaCount(6) = count6
        strFormulaElem(6) = elem6$
        intFormulaCount(7) = count7
        strFormulaElem(7) = elem7$
        intFormulaCount(8) = count8
        strFormulaElem(8) = elem8$
        intFormulaCount(9) = count9
        strFormulaElem(9) = elem9$
        intFormulaCount(10) = count10
        strFormulaElem(10) = elem10$
        
        ' Determine number of C, Si, N, P, O, S, Cl, I, F, Br and H atoms
        For X = 1 To 10
            Select Case strFormulaElem(X)
            Case "C": udtElementNum.C = udtElementNum.C + intFormulaCount(X)
            Case "Si": udtElementNum.Si = udtElementNum.Si + intFormulaCount(X)
            Case "N": udtElementNum.N = udtElementNum.N + intFormulaCount(X)
            Case "P": udtElementNum.P = udtElementNum.P + intFormulaCount(X)
            Case "O": udtElementNum.O = udtElementNum.O + intFormulaCount(X)
            Case "S": udtElementNum.S = udtElementNum.S + intFormulaCount(X)
            Case "Cl": udtElementNum.Cl = udtElementNum.Cl + intFormulaCount(X)
            Case "I": udtElementNum.I = udtElementNum.I + intFormulaCount(X)
            Case "F": udtElementNum.F = udtElementNum.F + intFormulaCount(X)
            Case "Br": udtElementNum.Br = udtElementNum.Br + intFormulaCount(X)
            Case "H": udtElementNum.H = udtElementNum.H + intFormulaCount(X)
            Case Else: udtElementNum.Other = udtElementNum.Other + intFormulaCount(X)
            End Select
        Next X
        
        ' Compute maximum number of hydrogens
        If udtElementNum.Si = 0 And udtElementNum.C = 0 And udtElementNum.N = 0 And _
           udtElementNum.P = 0 And udtElementNum.Other = 0 And (udtElementNum.O > 0 Or udtElementNum.S > 0) Then
            ' Only O and S
            udtElementNum.H = 3
        Else
            ' Formula is: [#C*2 + 3 - (2 if N or P present)] + [#N + 3 - (1 if C or Si present)] + [#other elements * 4 + 2], where we assume other elements can have a coordination number of up to 6
            If udtElementNum.C > 0 Or udtElementNum.Si > 0 Then udtElementNum.H = udtElementNum.H + (udtElementNum.C + udtElementNum.Si) * 2 + 3
            'If udtElementNum.Si > 0 Then udtElementNum.H = udtElementNum.H + udtElementNum.Si * 2 + 3
            If udtElementNum.N > 0 Or udtElementNum.P > 0 Then udtElementNum.H = udtElementNum.H + (udtElementNum.N + udtElementNum.P) + 3
            'If udtElementNum.P > 0 Then udtElementNum.H = udtElementNum.H + udtElementNum.P + 3
            If udtElementNum.Other > 0 Then udtElementNum.H = udtElementNum.H + udtElementNum.Other * 4 + 3
    
            ' Correction for carbon contribution
            'If (udtElementNum.C > 0 Or udtElementNum.Si > 0) And (udtElementNum.N > 0 Or udtElementNum.P > 0) Then udtElementNum.H = udtElementNum.H - 2
            ' Correction for nitrogen contribution
            'If (udtElementNum.N > 0 Or udtElementNum.P > 0) And (udtElementNum.C > 0 Or udtElementNum.Si > 0) Then udtElementNum.H = udtElementNum.H - 1
            ' Combine the above two if's into:
            If (udtElementNum.N > 0 Or udtElementNum.P > 0) And (udtElementNum.C > 0 Or udtElementNum.Si > 0) Then udtElementNum.H = udtElementNum.H - 3
        End If
        
        ' correct for if H only
        If udtElementNum.H < 3 Then udtElementNum.H = 3
        
        ' correct for halogens
        udtElementNum.H = udtElementNum.H - udtElementNum.F - udtElementNum.Cl - udtElementNum.Br - udtElementNum.I
        
        ' correct for negative udtElementNum.H
        If udtElementNum.H < 0 Then udtElementNum.H = 0
        
        ' Verify H's
        If udtElementNum.H <= udtElementNum.H Or Not cChkBox(frmFinderOptions.chkVerifyHydrogens) Then
            blnHOK = True
        Else
            blnHOK = False
        End If
        
        ' Only proceed if hydrogens check out
        If blnHOK Then
            
            ' See if sngTotalCharge is within charge limits
            If cChkBox(frmFinderOptions.chkFindCharge) Then
                sngTotalCharge = CorrectChargeEmpirical(sngTotalCharge, udtElementNum, blnChargeOK)
            Else
                blnChargeOK = 1
            End If
            
            'frmFinderOptions.chk
            ' If charge is within range and checking for multiples, see if correct m/z too
            If blnChargeOK And cChkBox(frmFinderOptions.chkFindTargetMtoZ) Then
                CheckMtoZWithTarget dblTotalWeight, sngTotalCharge, dblTargetWeight, dblWeightTolerance, intMultipleMtoZCharge, blnChargeOK
            End If
            
            If blnChargeOK Then
                ConstructAndVerifyCompound = True
            Else
                ConstructAndVerifyCompound = False
            End If
        End If
    Else
        ConstructAndVerifyCompound = True
    End If

    Exit Function
    
ConstructAndVerifyCompoundErrorHandler:
    GeneralErrorHandler "frmFinder|ConstructAndVerifyCompound", Err.Number, Err.Description
    ConstructAndVerifyCompound = False

End Function

Private Sub ConstructResultsLine(ByRef strResultsLine As String, dblTotalMass As Double, dblTargetMass As Double, sngTotalCharge As Single)
    ' dblTargetMass <= 0 means matching percent compositions, so don't want to add dm= to line
    
    Dim blnShowDeltaMass As Boolean
    
On Error GoTo ConstructResultsLineErrorHandler

    ' Add charge if desired
    If cChkBox(frmFinderOptions.chkFindCharge) Then
        strResultsLine = strResultsLine & "^" & Trim(CStr(sngTotalCharge))
    End If
    
    ' Add MW info and dm or ppm to strResultsLine
    If chkShowDeltaMass.value = vbChecked Then
        blnShowDeltaMass = True
    Else
        blnShowDeltaMass = False
    End If
    
    If blnShowDeltaMass And frmProgramPreferences.optStdDevType(1).value = True And dblTargetMass > 0 Then
        ' Use scientific notation
        If chkPpmMode.value = 0 Then
            strResultsLine = strResultsLine & vbTab & gMWAbbreviation & "=" & CStr(dblTotalMass) & _
                         vbTab & LookupLanguageCaption(10280, "dm") & "=" & _
                         Format$((dblTotalMass - dblTargetMass), "0.0#######E+00")
        Else
            strResultsLine = strResultsLine & vbTab & gMWAbbreviation & "=" & CStr(dblTotalMass) & _
                         vbTab & LookupLanguageCaption(10280, "dm") & "=" & _
                         Format$(CDbl(((dblTotalMass) / dblTargetMass - 1) * 1000000#), "0.0####E+00") & " " & LookupLanguageCaption(10285, "ppm")
        End If
        If frmFinderOptions.chkFindMtoZ.value = 1 And sngTotalCharge <> 0 Then
            ' Compute m/z value
            strResultsLine = strResultsLine & vbTab & "m/z " & _
                         Format$(Abs(dblTotalMass / sngTotalCharge), "0.0######E+00")
        End If
    Else
        If Not blnShowDeltaMass Or dblTargetMass <= 0 Then
            ' Don't add difference in mass (dm) result since standard deviation mode is off
            strResultsLine = strResultsLine & vbTab & gMWAbbreviation & "=" & CStr(dblTotalMass)
        Else
            ' Display dm result in 0.000 notation rather than exponential
            If chkPpmMode.value = 0 Then
                strResultsLine = strResultsLine & vbTab & gMWAbbreviation & "=" & CStr(dblTotalMass) & _
                             vbTab & LookupLanguageCaption(10280, "dm") & "=" & _
                             Format$((dblTotalMass - dblTargetMass), "#0.0#######")
            Else
                strResultsLine = strResultsLine & vbTab & gMWAbbreviation & "=" & CStr(dblTotalMass) & _
                             vbTab & LookupLanguageCaption(10280, "dm") & "=" & _
                             Format$(CDbl(((dblTotalMass) / dblTargetMass - 1) * 1000000#), "#0.0") & " " & LookupLanguageCaption(10285, "ppm")
            End If
        End If
        If frmFinderOptions.chkFindMtoZ.value = 1 And sngTotalCharge <> 0 Then
            ' Compute m/z value
            strResultsLine = strResultsLine & vbTab & "m/z " & _
                         Format$(Abs(dblTotalMass / sngTotalCharge), "#0.000")
        End If
    End If

    Exit Sub
    
ConstructResultsLineErrorHandler:
    GeneralErrorHandler "frmFinder|ConstructResultsLine", Err.Number, Err.Description

End Sub

Private Sub ConvertListToRTF()
    Dim X As Integer, Y As Integer, strWorkLine As String, strRTF As String, strRTFAdd As String
    
On Error GoTo ConvertListToRTFErrorHandler

    rtfResults.Text = ""
    
    strRTF = "{\rtf1\ansi\deff0\deftab720{\fonttbl{\f0\fswiss MS Sans Serif;}{\f1\froman\fcharset2 Symbol;}{\f2\froman " & objMwtWin.RtfFontName & ";}{\f3\froman Times New Roman;}{\f4\fswiss\fprq2 System;}}{\colortbl\red0\green0\blue0;\red255\green0\blue0;\red255\green255\blue255;}\deflang1033\pard\plain\f2\fs" & Trim(Str(CIntSafeDbl(objMwtWin.RtfFontSize * 2.25))) & " "
    
    For X = 0 To lstResults.ListCount - 1
        strWorkLine = lstResults.List(X)
        If optType(1).value = True And InStr(strWorkLine, LookupLanguageCaption(10535, "has")) > 0 Then
            ' Matching percent compositions and on a has ... line, just add to strRTF
            strRTF = strRTF & strWorkLine & "\par "
        Else
            ' Search for a Tab
            Y = InStr(strWorkLine, vbTab)
            If Y > 0 Then
                ' Found a tab, convert part before tab to formatted RTF
                strRTFAdd = ResultsToRtf(Left(strWorkLine, Y - 1)) & Mid(strWorkLine, Y)
            Else
                strRTFAdd = strWorkLine
            End If
            strRTF = strRTF & strRTFAdd & "\par "
        End If
        If X Mod 50 = 0 Then
            If lstResults.ListCount > 1 Then
                lblPercentCompleted.Caption = LookupLanguageCaption(10150, "Formatting") & " " & _
                                              Str(CIntSafeDbl(X / (lstResults.ListCount - 1) * 100)) & _
                                              LookupLanguageCaption(10110, "% Completed")
            End If
            DoEvents
            If gKeyPressAbortFormulaFinder >= 2 Then
                ' A key was pressed, so stop formatting
                If gKeyPressAbortFormulaFinder = 2 Then MsgBox LookupLanguageCaption(10450, "A key was pressed.") & "  " & LookupLanguageCaption(10465, "Stopping formatting."), _
                                            vbOKOnly + vbCritical, LookupLanguageCaption(10455, "Abort")
                If gKeyPressAbortFormulaFinder = 3 Then MsgBox LookupLanguageCaption(10460, "The mouse was clicked outside the results box.") & "  " & LookupLanguageCaption(10465, "Stopping formatting."), _
                                            vbOKOnly + vbCritical, LookupLanguageCaption(10455, "Abort")
                gKeyPressAbortFormulaFinder = 4
                Exit For
            End If
        End If
    Next X
    
    rtfResults.TextRTF = strRTF & "}"
    
    If gKeyPressAbortFormulaFinder < 2 Then
        lblPercentCompleted.Caption = LookupLanguageCaption(10155, "Done")
    Else
        lblPercentCompleted.Caption = LookupLanguageCaption(10160, "Formatting Aborted")
    End If
    
    Exit Sub

ConvertListToRTFErrorHandler:
    GeneralErrorHandler "frmFinder|ConvertListToRTF", Err.Number, Err.Description

End Sub

Private Sub CopyRoutinePrivate()
    ' Copy text from rtfResults to the clipboard
    
    Dim strCopyText As String
    
    strCopyText = rtfResults.SelText
        
    RemoveHeightAdjustChar strCopyText
    
    ' Copy corrected text to Clipboard.
    Clipboard.Clear
    Clipboard.SetText strCopyText, vbCFText

End Sub

Private Function CorrectChargeEmpirical(sngTotalCharge As Single, udtElementNum As udtElementNumType, blnChargeOK As Boolean) As Single

    Dim intNumHalogens As Integer, intIndex As Integer
    
    ' Uncomment to debug
    'If udtElementNum.C = 4 And udtElementNum.H = 4 And udtElementNum.N = 6 Then
    '    Debug.Print ("Debug: Check this")
    'End If
    
    ' Correct charge using rules for an empirical formula
    If udtElementNum.C + udtElementNum.Si >= 1 Then
        If udtElementNum.H > 0 And objMwtWin.GetElementStat(1, esCharge) = 1 Then
            ' Since carbon or silicon are present, assume the hydrogens should be negative
            ' Subtract udtElementNum.h*2 since hydrogen is assigned a +1 charge if ElementStats(1).Charge = 1
            sngTotalCharge = sngTotalCharge - udtElementNum.H * 2
        End If
        ' Correct for udtElementNum.ber of C and Si
        If udtElementNum.C + udtElementNum.Si > 1 Then
            sngTotalCharge = sngTotalCharge - CIntSafeDbl(udtElementNum.C + udtElementNum.Si - 1) * 2
        End If
    End If

    If udtElementNum.N + udtElementNum.P >= 1 And udtElementNum.C > 0 Then
        ' Assume 2 hydrogens around each Nitrogen or Phosphorus, thus add back +2 for each H
        ' First, decrease udtElementNum.ber of halogens by udtElementNum.ber of hydrogens & halogens taken up by the carbons
        ' Determine # of H taken up by all the carbons in a compound without N or P, then add back 1 H for each N and P
        intNumHalogens = udtElementNum.H + udtElementNum.F + udtElementNum.Cl + udtElementNum.Br + udtElementNum.I
        intNumHalogens = intNumHalogens - (udtElementNum.C * 2 + 2) + udtElementNum.N + udtElementNum.P
        
        If intNumHalogens >= 0 Then
            For intIndex = 1 To udtElementNum.N + udtElementNum.P
                sngTotalCharge = sngTotalCharge + 2
                intNumHalogens = intNumHalogens - 1
                If intNumHalogens <= 0 Then
                    Exit For
                Else
                    sngTotalCharge = sngTotalCharge + 2
                    intNumHalogens = intNumHalogens - 1
                    If intNumHalogens <= 0 Then Exit For
                End If
            Next intIndex
        End If
    End If

    If cChkBox(frmFinderOptions.chkLimitChargeRange) Then
        ' Make sure sngTotalCharge is within the specified range
        If sngTotalCharge >= CIntSafe(frmFinderOptions.txtChargeMin.Text) And _
           sngTotalCharge <= CIntSafe(frmFinderOptions.txtChargeMax.Text) Then
            ' Charge is within range
            blnChargeOK = True
        Else
            blnChargeOK = False
        End If
    Else
        blnChargeOK = True
    End If

    CorrectChargeEmpirical = sngTotalCharge
End Function

Private Sub DisplayIsoAbundanceForCurrent(Optional blnFromRtfResults As Boolean = False)
    Dim lngSelPosition As Long, lngMatchLoc As Long
    Dim intIndex As Integer
    
    Dim blnTabFound As Boolean
    
    Dim strCompareText As String, strCharge As String, strCompareTextCompacted As String
    Dim strWorking As String, strMW As String
    
    If lstResults.ListCount = 0 Then
        MsgBox LookupLanguageCaption(10550, "The results box is empty."), vbOKOnly, _
               LookupLanguageCaption(10555, "Nothing to Copy")
    Else
        ' Determine the line the cursor is on
        lngSelPosition = rtfResults.SelStart
        
        strWorking = rtfResults.Text
        
        If lngSelPosition = 0 Then lngSelPosition = 1
        
        ' Step backward until the previous vbCrLf is found
        Do While Mid(strWorking, lngSelPosition, 1) <> vbLf And lngSelPosition > 1
            lngSelPosition = lngSelPosition - 1
        Loop
        
        If lngSelPosition > 1 Then
            strWorking = Mid(strWorking, lngSelPosition + 1)
        Else
            ' At beginning of results; display first result
            If lstResults.ListCount > 1 Then
                strWorking = lstResults.List(1)
            End If
        End If
        
        ' Step forward until MW= is found (or foreign equivalent in gMWAbbreviation)
        strMW = gMWAbbreviation & "="
        
        lngSelPosition = 1
        blnTabFound = False
        Do While lngSelPosition < Len(strWorking)
            If Mid(strWorking, lngSelPosition, 1) = vbTab Then
                blnTabFound = True
                Exit Do
            End If
            lngSelPosition = lngSelPosition + 1
        Loop
        
        If blnTabFound Then
            strWorking = Left(strWorking, lngSelPosition - 1)
            
            ' Replace any tilde character in strWorking
            strWorking = Trim(Replace(strWorking, RTF_HEIGHT_ADJUSTCHAR, ""))
            
            ' If a Caret is present, remove the caret and any numbers following
            lngMatchLoc = InStr(strWorking, "^")
            If lngMatchLoc > 0 Then
                strWorking = Left(strWorking, lngMatchLoc - 1)
            Else
                ' Need to look for entry in lstResults in order to return just the formula, and not the charge
                For intIndex = 0 To lstResults.ListCount - 1
                    strCompareText = lstResults.List(intIndex)
                    lngMatchLoc = InStr(strCompareText, vbTab)
                    If lngMatchLoc > 0 Then strCompareText = Left(strCompareText, lngMatchLoc - 1)
                    lngMatchLoc = InStr(strCompareText, "^")
                    If lngMatchLoc > 0 Then
                        ' Need to reverse the charge (i.e. change from -3 or +3 to 3- or 3+
                        ' Also need to remove the ^ from strCompareText
                        strCharge = Trim(Mid(strCompareText, lngMatchLoc + 1))
                        strCompareText = Trim(Left(strCompareText, lngMatchLoc - 1))
                        If Left(strCharge, 1) = "+" Then
                            strCharge = Mid(strCharge, 2) & "+"
                        ElseIf Left(strCharge, 1) = "-" Then
                            strCharge = Mid(strCharge, 2) & "-"
                        Else
                            strCharge = strCharge & "+"
                        End If
                        strCompareTextCompacted = strCompareText & strCharge
                    Else
                        strCompareTextCompacted = strCompareText
                    End If
                    If strCompareTextCompacted = strWorking Then
                        strWorking = strCompareText
                        Exit For
                    End If
                Next intIndex
            End If
            
            Me.MousePointer = vbHourglass
            
            With frmIsotopicDistribution
                .rtfFormula = strWorking
                SetCheckBox .chkPlotResults, True
            
                .SetNeedToZoomOutFull True
                .SetFormActivated
                
                .StartIsotopicDistributionCalcs False
            End With
            
            Me.MousePointer = vbNormal
            
            If blnFromRtfResults Then
                rtfResults.SetFocus
            End If
        Else
            MsgBox "The cursor is not on a line with a valid formula."
        End If
        
    End If
End Sub

Private Function FinderMatchStats(matchline As Integer, Optional strMatchSymbol As String, Optional Charge As Single) As Double
    ' Returns the weight of the element on matchline, along with setting the symbol and charge
    Dim intSymbolReference As Integer, strWork As String
    Dim blnSymbolMatched As Boolean
    Dim strCustomElementPhrase As String, strNumOrElementPhrase As String
    Dim strSymbol As String, dblMass As Double
    Dim lngError As Long
    Dim sngCharge As Single
    Dim strFormula As String
    Dim blnIsAminoAcid As Boolean
    
    strCustomElementPhrase = LookupLanguageCaption(10420, "Custom")
    strNumOrElementPhrase = LookupLanguageCaption(10270, "# or Element or Abbrev.")
    
    Select Case matchline
        Case 0 To 3
            Select Case matchline
            Case 0: intSymbolReference = 6
            Case 1: intSymbolReference = 1
            Case 2: intSymbolReference = 7
            Case 3: intSymbolReference = 8
            End Select
            
            lngError = objMwtWin.GetElement(intSymbolReference, strSymbol, dblMass, 0, sngCharge, 0)
            Debug.Assert lngError = 0
            
            FinderMatchStats = dblMass
            strMatchSymbol = strSymbol
            Charge = sngCharge
        
        Case 4 To 9:
            ' Working with custom elements
            If Not IsNumeric(txtWeight(matchline).Text) Then
                blnSymbolMatched = False
                If txtWeight(matchline).Text = strNumOrElementPhrase Then
                    ' The custom element phrase is present; don't try to match it to an element
                Else
                    ' A single element or abbreviation was entered
                    
                    ' Convert input to default format of first letter capitalized and rest lowercase
                    txtWeight(matchline).Text = UCase(Left(txtWeight(matchline).Text, 1)) + _
                                                Mid(txtWeight(matchline).Text, 2)
                    strWork = txtWeight(matchline).Text
                        
                    ' See if this is an element
                    intSymbolReference = objMwtWin.GetElementID(strWork)
                    
                    If intSymbolReference >= 1 Then
                        ' Found an element
                        lngError = objMwtWin.GetElement(intSymbolReference, strSymbol, dblMass, 0, sngCharge, 0)
                        Debug.Assert lngError = 0
                        
                        FinderMatchStats = dblMass
                        strMatchSymbol = strSymbol
                        Charge = sngCharge
                        
                        blnSymbolMatched = True
                    Else
                        ' See if this is an abbreviation
                        intSymbolReference = objMwtWin.GetAbbreviationID(strWork)
                        If intSymbolReference >= 1 Then
                            ' Found a normal abbreviation
                            lngError = objMwtWin.GetAbbreviation(intSymbolReference, strSymbol, strFormula, sngCharge, blnIsAminoAcid)
                            FinderMatchStats = objMwtWin.ComputeMass(strFormula)
                            strMatchSymbol = strSymbol
                            Charge = sngCharge
                            
                            blnSymbolMatched = True
                        End If
                    End If
                    
                End If
                If Not blnSymbolMatched Then
                    FinderMatchStats = -1
                    txtWeight(matchline) = strNumOrElementPhrase
                    strMatchSymbol = Left(strCustomElementPhrase, 1) & Trim(Str(matchline - 3)) & "_"
                    Charge = 0
                Else
                    ' Re-assign matched element or abbreviation to properly capitalize
                    txtWeight(matchline) = strMatchSymbol
                End If
            Else
                FinderMatchStats = CDblSafe(txtWeight(matchline))
                strMatchSymbol = Left(strCustomElementPhrase, 1) & Trim(Str(matchline - 3)) & "_"
                Charge = 0
            End If
        Case Else:
            FinderMatchStats = -1
    End Select

End Function

Public Function FormulaFinderCalculate(Optional blnQuietMode As Boolean = False) As String
    
    Dim intNumPotentialElementCount As Integer, blnCalculationsAborted As Boolean
    Dim X As Integer, Y As Integer, intMaxHits As Integer, eResponse As VbMsgBoxResult
    Dim blnBad As Boolean, lngCount As Long
    Dim intMultipleSearchMin As Integer, intMultipleSearchMax As Integer
    Dim tolerance As Double, dblTargetWeight As Double, percentcompsum As Double
    Dim strMessage As String, strWork As String, strSwap As String
    Dim dblSwapVal As Double
    Dim strMatchSymbol As String, MaxFormulaWeight As Double, Charge As Single
    
    Static blnCalculating As Boolean
    
    Dim intPotentialElementPointers(0) As Integer              ' Empty array of pointers to the potential elements
    
    Dim intRange(MAX_MATCHINGELEMENTS, 2) As Integer        ' The min and max number for each element
    Dim dblPotentialElementStats(MAX_MATCHINGELEMENTS, 1) As Double          ' The elemental weight and charge for each element or abbreviation
    Dim strPotentialElements(MAX_MATCHINGELEMENTS) As String               ' The symbol for each element or abbreviation
    Dim dblTargetPercents(MAX_MATCHINGELEMENTS, 1) As Double         ' The min and max percentages for each element (if applicable)
        
    Dim strCompoundList() As String
    Dim strCodeString() As String
    Dim udtResultStats() As udtResultStatsType
    Dim intPointerArray() As Integer
        
    If blnCalculating Then Exit Function
    
    gKeyPressAbortFormulaFinder = 0
    blnCalculationsAborted = False
    
On Error GoTo FinderErrorHandler
    
    ' Check for Checked elements
    strMessage = ""
    For X = 0 To MAX_MATCHINGELEMENTS
        If cChkBox(chkElements(X)) Then Exit For
    Next X
           
    If X = MAX_MATCHINGELEMENTS + 1 Then
        strMessage = strMessage & LookupMessage(700) & vbCrLf
    Else
        ' Check for valid information in text boxes
        If Val(txtHits.Text) <= 0 Then
            strMessage = strMessage & LookupMessage(705) & vbCrLf
        ElseIf Val(txtHits.Text) > MAX_HITS Then
            strMessage = strMessage & LookupMessage(710, " " & Str(MAX_HITS)) & vbCrLf
        Else
            intMaxHits = CIntSafe(txtHits.Text)
        End If
        
        ' Check for reasonable numbers in Min/Max boxes
        For X = 0 To 9
            If cChkBox(chkElements(X)) And Val(txtMin(X).Text) < 0 Then
                strMessage = strMessage & LookupMessage(715) & vbCrLf
            ElseIf cChkBox(chkElements(X)) And Val(txtMin(X).Text) > Val(txtMax(X).Text) Then
                strMessage = strMessage & LookupMessage(720) & vbCrLf
            ElseIf cChkBox(chkElements(X)) And Val(txtMax(X).Text) > 65025 Then
                strMessage = strMessage & LookupMessage(725) & vbCrLf
            End If
        Next X
                
        ' Initialize custom elements
        For X = 4 To MAX_MATCHINGELEMENTS - 1
            If cChkBox(chkElements(X)) Then
                If Not IsNumeric(txtWeight(X).Text) Then
                    If txtWeight(X).Text = LookupLanguageCaption(10270, "# or Element or Abbrev.") Then strMessage = strMessage & LookupMessage(730) & vbCrLf
                Else
                    If Val(txtWeight(X).Text) <= 0 Then strMessage = strMessage & LookupMessage(735) & vbCrLf
                End If
            End If
        Next X
        
        If optType(0).value = True Then
            ' Matching Molecular Weights
            If Not IsNumeric(txtMWT.Text) Then
                strMessage = strMessage & LookupMessage(740) & vbCrLf
            Else
                If Val(txtMWT.Text) <= 0 Then strMessage = strMessage & LookupMessage(745) & vbCrLf
            End If
            If Val(txtWeightTolerance.Text) < 0 Then strMessage = strMessage & LookupMessage(750) & vbCrLf
        Else
            ' Matching Percent Compositions
            
            ' Confirm that txtPercentMaxWeight is Valid
            If Not IsNumeric(txtPercentMaxWeight.Text) Then
                strMessage = strMessage & LookupMessage(755) & vbCrLf
            Else
                If Val(txtPercentMaxWeight.Text) <= 0 Then strMessage = strMessage & LookupMessage(760) & vbCrLf
            End If
            MaxFormulaWeight = CDblSafe(txtPercentMaxWeight.Text)
            
            For X = 0 To 9
                If cChkBox(chkElements(X)) Then
                    If Not IsNumeric(txtPercent(X).Text) Then
                        FinderMatchStats X, strMatchSymbol
                        strMessage = strMessage & LookupMessage(765, ": " & strMatchSymbol) & vbCrLf
                    Else
                        If Val(txtPercent(X).Text) <= 0 Then strMessage = strMessage & LookupMessage(770) & vbCrLf
                        percentcompsum = percentcompsum + CDblSafe(txtPercent(X).Text)
                    End If
                End If
            Next X
            If strMessage = "" Then
                If percentcompsum > 100 + CDblSafe(txtPercentTolerance.Text) Or _
                   percentcompsum < 100 - CDblSafe(txtPercentTolerance.Text) Then
                    strMessage = LookupLanguageCaption(10480, "The sum of the percent compositions is not 100%") & " (" & CStr(percentcompsum) & "%)  " & LookupLanguageCaption(10485, "Continue with calculation?")
                    eResponse = MsgBox(strMessage, vbYesNo + vbQuestion + vbDefaultButton2, LookupLanguageCaption(4050, "Caution"))
                    If eResponse = vbNo Then
                        Exit Function
                    Else
                        strMessage = ""
                    End If
                End If
            End If
        End If
    End If

    If strMessage <> "" Then
        If Not blnQuietMode Then
            MsgBox strMessage, vbOKOnly + vbExclamation, LookupLanguageCaption(10490, "Cannot Calculate")
        Else
            FormulaFinderCalculate = strMessage
        End If
        Exit Function
    End If
    
    ' Delete old Results
    lstResults.Clear
    lstResults.Visible = True
    rtfResults.Visible = False
    
    ' Fill working variables with correct values
    If optType(1).value = True Then
        ' Matching Percent Compositions
        tolerance = CDblSafe(txtPercentTolerance.Text)
    Else
        ' Matching Molecular Weights
        If chkPpmMode.value = 0 Then
            tolerance = CDblSafe(txtWeightTolerance.Text)
        Else
            tolerance = Abs(CDblSafe(txtMWT.Text) * CDblSafe(txtWeightTolerance.Text) / 1000000#)
        End If
        dblTargetWeight = CDblSafe(txtMWT.Text)
    End If
    
    ' Determine which elements are checked and add to potentialelement list
    intNumPotentialElementCount = 0
    For X = 0 To 9
        If cChkBox(chkElements(X)) Then
            intRange(intNumPotentialElementCount, 0) = CIntSafe(txtMin(X).Text)
            intRange(intNumPotentialElementCount, 1) = CIntSafe(txtMax(X).Text)
            If X >= 4 Then
                ' Working with custom elements
                If Not IsNumeric(txtWeight(X).Text) Then
                    ' A single element or abbreviation was entered
                    blnBad = False
                    ' Convert input to default format of first letter capitalized and rest lowercase
                    txtWeight(X).Text = UCase(Left(txtWeight(X).Text, 1)) + _
                                        Mid(txtWeight(X).Text, 2)
                    strWork = txtWeight(X).Text
                    For Y = 1 To Len(strWork)
                        If (Asc(UCase(Mid(strWork, Y, 1))) < 65 Or Asc(UCase(Mid(strWork, Y, 1))) > 90) And _
                           Asc(UCase(Mid(strWork, Y, 1))) <> 43 And _
                           Asc(UCase(Mid(strWork, Y, 1))) <> 95 Then
                            blnBad = True
                            If Not blnQuietMode Then
                                MsgBox LookupMessage(775), vbOKOnly + vbExclamation, LookupMessage(350)
                            Else
                                FormulaFinderCalculate = LookupMessage(775)
                            End If
                            
                            Exit Function
                        End If
                    Next Y

                    If Len(txtWeight(X).Text) = 0 Then
                        ' Too short
                        blnBad = True
                        If Not blnQuietMode Then
                            MsgBox LookupMessage(780), vbOKOnly + vbExclamation, LookupMessage(350)
                        Else
                            FormulaFinderCalculate = LookupMessage(780)
                        End If
                        
                    Else
                        Charge = 0
                        dblPotentialElementStats(intNumPotentialElementCount, 0) = FinderMatchStats(X, strMatchSymbol, Charge)     ' Returns weight of element/abbreviation, but also charge
                        dblPotentialElementStats(intNumPotentialElementCount, 1) = Charge
                        
                        If dblPotentialElementStats(intNumPotentialElementCount, 0) = -1 Then
                            ' Element or abbreviation not found
                            blnBad = True
                            If Not blnQuietMode Then
                                MsgBox LookupMessage(785) & ": " & strWork & "  " & LookupMessage(790), vbOKOnly + vbExclamation, LookupMessage(350)
                            Else
                                FormulaFinderCalculate = LookupMessage(785) & ": " & strWork
                            End If
                        Else
                            ' No problems, store symbol
                            strPotentialElements(intNumPotentialElementCount) = strMatchSymbol
                        End If
                    End If
                Else
                    ' Custom element, only weight given so charge is 0
                    dblPotentialElementStats(intNumPotentialElementCount, 0) = CDblSafe(txtWeight(X))
                    dblPotentialElementStats(intNumPotentialElementCount, 1) = 0
                    strPotentialElements(intNumPotentialElementCount) = Left(LookupLanguageCaption(10420, "Custom"), 1) & Trim(Str(X - 3)) & "_"
                End If
            Else
                dblPotentialElementStats(intNumPotentialElementCount, 0) = FinderMatchStats(X, strMatchSymbol, Charge)
                dblPotentialElementStats(intNumPotentialElementCount, 1) = Charge
                strPotentialElements(intNumPotentialElementCount) = strMatchSymbol
            End If
            dblTargetPercents(intNumPotentialElementCount, 0) = CDblSafe(txtPercent(X).Text) - tolerance  ' Lower bound of target percentage
            dblTargetPercents(intNumPotentialElementCount, 1) = CDblSafe(txtPercent(X).Text) + tolerance  ' Upper bound of target percentage
            intNumPotentialElementCount = intNumPotentialElementCount + 1
        End If
        If blnBad Then Exit Function
    Next X
    
    ' Change mouse pointer to hourglass with arrow, so user can still multitask
    MousePointer = vbArrowHourglass
    
    ' Set the key-watch variable to one
    gKeyPressAbortFormulaFinder = 1
    blnCalculating = True
    cmdCalculate.Visible = False
    cmdOK.Visible = False
    cmdPrint.Visible = False
    cmdDisplayIsotopicAbundance.Visible = False
    
    ' Make the percent completed box visible
    lblPercentCompleted.Caption = LookupLanguageCaption(10120, "Searching") & " 0% " & LookupLanguageCaption(10130, "Completed")
    lblPercentCompleted.Visible = True
    
    If frmFinderOptions.cboSearchType.ListIndex = 0 Then
        ' Thorough search
        
        ' Reorder dblPotentialElementStats pointer array in order from heaviest to lightest element
        ' Greatly speeds up the recursive routine
        ' Bubble sort
        For Y = intNumPotentialElementCount - 1 To 1 Step -1       ' Sort from end to start
            For X = 0 To Y - 1
                If dblPotentialElementStats(X, 0) < dblPotentialElementStats(X + 1, 0) Then
                    ' Swap the element symbols
                    strSwap = strPotentialElements(X)
                    strPotentialElements(X) = strPotentialElements(X + 1)
                    strPotentialElements(X + 1) = strSwap
                    ' and their weights
                    dblSwapVal = dblPotentialElementStats(X, 0)
                    dblPotentialElementStats(X, 0) = dblPotentialElementStats(X + 1, 0)
                    dblPotentialElementStats(X + 1, 0) = dblSwapVal
                    ' and their charge
                    dblSwapVal = dblPotentialElementStats(X, 1)
                    dblPotentialElementStats(X, 1) = dblPotentialElementStats(X + 1, 1)
                    dblPotentialElementStats(X + 1, 1) = dblSwapVal
                    
                    If optType(1).value = True Then
                        ' and the dblTargetPercents array
                        dblSwapVal = dblTargetPercents(X, 0)
                        dblTargetPercents(X, 0) = dblTargetPercents(X + 1, 0)
                        dblTargetPercents(X + 1, 0) = dblSwapVal
                        
                        dblSwapVal = dblTargetPercents(X, 1)
                        dblTargetPercents(X, 1) = dblTargetPercents(X + 1, 1)
                        dblTargetPercents(X + 1, 1) = dblSwapVal
                        
                    End If
                End If
            Next X
        Next Y
        
        ' Initialize the status box
        UpdateStatus intNumPotentialElementCount, 3, True, 0
        
        If optType(0).value = True Then
            If cChkBox(frmFinderOptions.chkFindTargetMtoZ) Then
                ' Searching for target m/z rather than target mass
                MultipleSearchMath intNumPotentialElementCount, intMultipleSearchMin, intMultipleSearchMax
                
                For X = intMultipleSearchMin To intMultipleSearchMax
                    ' Call the RecursiveMWfinder repeatedly, sending dblTargetWeight * x each time to search for target, target*2, target*3, etc.
                    RecursiveMWFinder strPotentialElements(), dblPotentialElementStats(), 0, intNumPotentialElementCount, intPotentialElementPointers(), 0, 0, dblTargetWeight * X, tolerance, 0, X
                Next X
            Else
                RecursiveMWFinder strPotentialElements(), dblPotentialElementStats(), 0, intNumPotentialElementCount, intPotentialElementPointers(), 0, 0, dblTargetWeight, tolerance, 0, 0
            End If
        Else
            RecursivePCompFinder strPotentialElements(), dblPotentialElementStats(), 0, intNumPotentialElementCount, intPotentialElementPointers(), 0, 0, dblTargetPercents(), MaxFormulaWeight, 9
        End If
        
    Else
        ' Bounded search
        
        If cChkBox(frmFinderOptions.chkFindTargetMtoZ) Then
            ' Searching for target m/z rather than target mass
            
            MultipleSearchMath intNumPotentialElementCount, intMultipleSearchMin, intMultipleSearchMax
            ' Send max and min charge for searching
            
            OldFormulaFinder strPotentialElements(), dblPotentialElementStats(), intNumPotentialElementCount, dblTargetWeight, tolerance, intMultipleSearchMin, intMultipleSearchMax, intRange(), MaxFormulaWeight, intMaxHits, dblTargetPercents()
        Else
            OldFormulaFinder strPotentialElements(), dblPotentialElementStats(), intNumPotentialElementCount, dblTargetWeight, tolerance, 1, 1, intRange(), MaxFormulaWeight, intMaxHits, dblTargetPercents()
        End If
    End If
    
    ' Show abort messages if necessary
    If gKeyPressAbortFormulaFinder >= 2 Then
        If gKeyPressAbortFormulaFinder = 2 Then MsgBox LookupLanguageCaption(10500, "A key was pressed.") & "  " & _
                                    LookupLanguageCaption(10510, "Stopping calculations."), vbOKOnly + vbCritical, _
                                    LookupLanguageCaption(10455, "Abort")
        If gKeyPressAbortFormulaFinder = 3 Then MsgBox LookupLanguageCaption(10505, "The mouse was clicked outside the results box.") & "  " & _
                                    LookupLanguageCaption(10510, "Stopping calculations."), vbOKOnly + vbCritical, _
                                    LookupLanguageCaption(10455, "Abort")
        gKeyPressAbortFormulaFinder = 1
        blnCalculationsAborted = True
    ElseIf lstResults.ListCount >= intMaxHits Then
        strMessage = LookupLanguageCaption(10515, "The maximum number of hits has been reached.") & "  " & LookupLanguageCaption(10510, "Stopping calculations.")
        If Not blnQuietMode Then
            MsgBox strMessage, vbOKOnly + vbCritical, LookupLanguageCaption(10520, "Maximum Hits")
        End If
    End If

    ' Determine total number of results found
    lngCount = Val(lstResults.ListCount)
    
    ' Sort results if necessary (lngCount from above is used for Sorting... line and after sorting)
    If cChkBox(frmFinderOptions.chkSort) Then  ' And gKeyPressAbortFormulaFinder < 2 Then
        ' Sort results
        lblPercentCompleted.Caption = LookupLanguageCaption(10115, "Sorting") & _
                                      Str(lngCount) & " " & _
                                      LookupLanguageCaption(10135, "compounds")
        
        ' Change view of lstresults box so first line is on top, avoids screen update problem when sorting
        If lstResults.ListCount > 0 Then
            lstResults.ListIndex = 0
            lstResults.ListIndex = -1
        End If
        
        ' Check for keypress and update the status box
        DoEvents

        ReDim strCodeString(lstResults.ListCount + 1)
        ReDim udtResultStats(lstResults.ListCount + 1)
        ReDim intPointerArray(lstResults.ListCount + 1)
        
        ' Convert compounds in list to code
        CompoundToCode dblTargetWeight, strCodeString(), udtResultStats(), intPointerArray()
        
        ' Sort the compounds
        ShellSortResults strCodeString(), udtResultStats(), intPointerArray(), 0, lstResults.ListCount - 1
        
        ReDim strCompoundList(lstResults.ListCount)            ' Temporary storage for the results box
        
        ' First, copy all of the results to a temporary array (I know it eats up memory, but I have no choice)
        For X = 0 To lstResults.ListCount - 1
            strCompoundList(X) = lstResults.List(X)
        Next X
        
        ' Now, put them back into the lstResults.ListCount box in the correct order
        ' Use intPointerArray() for this
        For X = 0 To lstResults.ListCount - 1
            lstResults.List(X) = strCompoundList(intPointerArray(X))
        Next X
        
        If gKeyPressAbortFormulaFinder < 2 Then
            lblPercentCompleted.Caption = "100% " & LookupLanguageCaption(10130, "Completed")
        Else
            lblPercentCompleted.Caption = LookupLanguageCaption(10140, "Sorting Interrupted")
        End If
    Else
        ' Don't sort
        If gKeyPressAbortFormulaFinder < 2 Then
            lblPercentCompleted.Caption = "100% " & LookupLanguageCaption(10130, "Completed")
        Else
            lblPercentCompleted.Caption = LookupLanguageCaption(10145, "Calculations Interrupted")
        End If
    End If
    
    If optType(1).value = True Then
        ' Matching percent compositions
        ' Separate results into two lines per compound for better readability
        
        ' First, change view of lstresults box so first line is on top, avoids screen update problem when sorting
        If lstResults.ListCount > 0 Then
            lstResults.ListIndex = 0
            lstResults.ListIndex = -1
        End If
        X = 0
        Do While X <= lstResults.ListCount - 1
            Y = InStr(lstResults.List(X), " " & LookupLanguageCaption(10535, "has"))
            If Y > 0 Then
                lstResults.AddItem Mid(lstResults.List(X), Y), X + 1
                lstResults.List(X) = Left(lstResults.List(X), Y - 1)
                X = X + 1
            End If
            X = X + 1
        Loop
    End If
        
    ' Add line for total number of compounds found
    strMessage = LookupLanguageCaption(10530, "Compounds found") & ": " & Str(lngCount)
    lstResults.AddItem strMessage, 0
    
    ' Set the ToolTipText for the listing
    If lngCount > 0 And cChkBox(frmProgramPreferences.chkShowToolTips) Then
        lstResults.ToolTipText = LookupToolTipLanguageCaption(10200, "Double click any line to expand it")
    Else
        lstResults.ToolTipText = ""
    End If
    
    blnCalculating = False

    ' Copy results from lstResults to rtfResults
    If gKeyPressAbortFormulaFinder < 2 Then
        ConvertListToRTF
        lstResults.Visible = False
        rtfResults.Visible = True
    End If
    
    If blnCalculationsAborted Then
        lblPercentCompleted.Caption = LookupLanguageCaption(10145, "Calculations Interrupted")
    End If
    

    ' Make sure results fit
    Form_Resize
    
FinderResume:
    
    ' Reset gKeyPressAbortFormulaFinder
    gKeyPressAbortFormulaFinder = 0

    cmdCalculate.Visible = True
    cmdOK.Visible = True
    cmdPrint.Visible = True
    cmdDisplayIsotopicAbundance.Visible = True
    
    ' Change mouse pointer to default
    MousePointer = vbDefault
    
    Exit Function

FinderErrorHandler:
    ' Change mouse pointer to default
    MousePointer = vbDefault
    
    If Err.Number = 7 Then
        MsgBox LookupLanguageCaption(10540, "A memory error has occurred.  The results list is probably full.  Results so far are shown."), _
                vbOKOnly + vbExclamation, LookupLanguageCaption(10545, "Out of Memory")
    Else
        GeneralErrorHandler "frmFinder|cmdCalculate", Err.Number, Err.Description
    End If
    Resume FinderResume

End Function

Private Sub RecursiveMWFinder(strPotentialElements() As String, dblPotentialElementStats() As Double, intStartIndex As Integer, intNumPotentialElementCount As Integer, intPotentialElementPointers() As Integer, intPointerCount As Integer, dblPotentialMassTotal As Double, dblTargetMass As Double, dblMassTolerance As Double, sngPotentialChargeTotal As Single, intMultipleMtoZCharge As Integer)
    ' Potentialelem contains the weights of all of the potential elements
    ' intPotentialElementPointers contains pointers to the elements that have been added to the potential formula so far
    ' potentialweight contains the weight of the potential formula
    ' if intMultipleMtoZCharge = 0 then not performing a multiple m/z search
    
    Dim intNewPotentialElementPointers() As Integer
    Dim intCurrentIndex As Integer, intPointer As Integer, dblNewMassTotal As Double
    Dim sngChargeTotal As Single, intExtra As Integer
                    
On Error GoTo RecursiveMWFinderErrorHandler

    ReDim intNewPotentialElementPointers(intPointerCount + 1)
                    
    If gKeyPressAbortFormulaFinder >= 2 Or lstResults.ListCount >= CIntSafe(txtHits.Text) Then
        Exit Sub
    End If
    
    For intCurrentIndex = intStartIndex To intNumPotentialElementCount - 1  ' intNumPotentialElementCount >= 1, if 1, means just dblPotentialElementStats(0, 0), etc.
        dblNewMassTotal = dblPotentialMassTotal + dblPotentialElementStats(intCurrentIndex, 0)
        sngChargeTotal = sngPotentialChargeTotal + dblPotentialElementStats(intCurrentIndex, 1)
        If dblNewMassTotal <= dblTargetMass + dblMassTolerance Then
            ' Below or within dblMassTolerance, add current element's pointer to pointer array
            For intPointer = 0 To intPointerCount - 1
                intNewPotentialElementPointers(intPointer) = intPotentialElementPointers(intPointer)
            Next intPointer
            intNewPotentialElementPointers(intPointerCount) = intCurrentIndex   ' Store current element's number at end of intNewPotentialElementPointers array
'            ReportCompound "", strPotentialElements(), dblPotentialElementStats(), intNumPotentialElementCount, intNewPotentialElementPointers(), intPointerCount + 1, dblNewMassTotal, dblTargetMass, dblMassTolerance, sngChargeTotal, intMultipleMtoZCharge
            
            ' Update status
            lngRecursiveFunctionCallCount = lngRecursiveFunctionCallCount + 1
            If lngRecursiveFunctionCallCount Mod 500 = 0 Then
                DoEvents
                If gKeyPressAbortFormulaFinder >= 2 Or lstResults.ListCount >= CIntSafe(txtHits.Text) Then
                    Exit Sub
                End If
            End If
            If intPointerCount = 3 Then UpdateStatus intNumPotentialElementCount, intPointerCount
                        
            If dblNewMassTotal >= dblTargetMass - dblMassTolerance Then
                ' Within dblMassTolerance, so verify hydrogens and add to lstResults
                RecursiveConstructAndVerifyCompound "", strPotentialElements(), dblPotentialElementStats(), intNumPotentialElementCount, intNewPotentialElementPointers(), intPointerCount + 1, dblNewMassTotal, dblTargetMass, dblMassTolerance, sngChargeTotal, intMultipleMtoZCharge
            End If
            
            ' Haven't reached dblTargetMass - dblMassTolerance region, so call RecursiveFinder again
            If intCurrentIndex = intNumPotentialElementCount - 1 Then
                ' But first, if adding the lightest element (i.e. the last in the list),
                ' add a bunch of it until the potential compound's weight is close to the target
                intExtra = 0
                Do While dblNewMassTotal < _
                         dblTargetMass - dblMassTolerance - dblPotentialElementStats(intCurrentIndex, 0)
                    intExtra = intExtra + 1
                    dblNewMassTotal = dblNewMassTotal + dblPotentialElementStats(intCurrentIndex, 0)
                    sngChargeTotal = sngChargeTotal + dblPotentialElementStats(intCurrentIndex, 1)
                Loop
                If intExtra > 0 Then
                    ReDim Preserve intNewPotentialElementPointers(intPointerCount + 1 + intExtra)
                    For intPointer = intPointerCount + 1 To intPointerCount + intExtra
                        intNewPotentialElementPointers(intPointer) = intCurrentIndex
                    Next intPointer
                    intPointerCount = intPointerCount + intExtra
                End If
            End If
            ' Now recursively call this sub
            RecursiveMWFinder strPotentialElements(), dblPotentialElementStats(), intCurrentIndex, intNumPotentialElementCount, intNewPotentialElementPointers(), intPointerCount + 1, dblNewMassTotal, dblTargetMass, dblMassTolerance, sngChargeTotal, intMultipleMtoZCharge
        End If
    Next intCurrentIndex

    Exit Sub

RecursiveMWFinderErrorHandler:
    GeneralErrorHandler "frmFinder|RecursiveMWFinder", Err.Number, Err.Description

End Sub

Private Sub RecursivePCompFinder(strPotentialElements() As String, dblPotentialElementStats() As Double, intStartIndex As Integer, intNumPotentialElementCount As Integer, intPotentialElementPointers() As Integer, intPointerCount As Integer, dblPotentialMassTotal As Double, dblTargetPercents() As Double, dblMaxFormulaWeight As Double, sngPotentialChargeTotal As Single)
    ' dblPotentialElementStats contains the weights of all of the potential elements
    ' intPotentialElementPointers contains pointers to the elements that have been added to the potential formula so far
    ' dblPotentialMassTotal contains the mass of the potential formula
    
    ' dblTargetPercents holds the lower and upper bounds of target percentages
    
    Dim intNewPotentialElementPointers() As Integer
    Dim dblPotentialPercents() As Double
    
    Dim intElementCount(MAX_MATCHINGELEMENTS) As Integer
    
    Dim intCurrentIndex As Integer, intPointer As Integer, dblNewMassTotal As Double
    Dim intPercentTrack As Integer
    Dim intIndex As Integer, sngChargeTotal As Single
    
On Error GoTo RecursivePCompFinderErrorHandler

    ReDim intNewPotentialElementPointers(intPointerCount + 1)
    ReDim dblPotentialPercents(intNumPotentialElementCount)
                    
    If gKeyPressAbortFormulaFinder >= 2 Or lstResults.ListCount >= CIntSafe(txtHits.Text) Then
        Exit Sub
    End If
    
    For intCurrentIndex = intStartIndex To intNumPotentialElementCount - 1  ' intNumPotentialElementCount >= 1, if 1, means just dblPotentialElementStats(0,0), etc.
        dblNewMassTotal = dblPotentialMassTotal + dblPotentialElementStats(intCurrentIndex, 0)
        sngChargeTotal = sngPotentialChargeTotal + dblPotentialElementStats(intCurrentIndex, 1)
        
        If dblNewMassTotal <= dblMaxFormulaWeight Then
            ' only proceed if weight is less than max weight
                
            ' add current element's pointer to pointer array
            For intPointer = 0 To intPointerCount - 1
                intNewPotentialElementPointers(intPointer) = intPotentialElementPointers(intPointer)
            Next intPointer
            intNewPotentialElementPointers(intPointerCount) = intCurrentIndex   ' Store current element's number at end of intNewPotentialElementPointers array
                
            Erase intElementCount()
            
            ' Compute the number of each element
            For intIndex = 0 To intPointerCount
                ' For each pointer, increment intElementCount
                intElementCount(intNewPotentialElementPointers(intIndex)) = intElementCount(intNewPotentialElementPointers(intIndex)) + 1
            Next intIndex
            
            For intIndex = 0 To intNumPotentialElementCount - 1
                If intElementCount(intIndex) = 0 Then Exit For
            Next intIndex
            
            If intIndex = intNumPotentialElementCount Then
                ' Only proceed if all elements are present
                ' Compute % comp of each element
                For intIndex = 0 To intNumPotentialElementCount - 1
                    dblPotentialPercents(intIndex) = intElementCount(intIndex) * dblPotentialElementStats(intIndex, 0) / dblNewMassTotal * 100
                Next intIndex
                'If intPointerCount = 0 Then dblPotentialPercents(0) = 100
        
                intPercentTrack = 0
                For intIndex = 0 To intNumPotentialElementCount - 1
                    If dblPotentialPercents(intIndex) >= dblTargetPercents(intIndex, 0) And _
                       dblPotentialPercents(intIndex) <= dblTargetPercents(intIndex, 1) Then
                        intPercentTrack = intPercentTrack + 1
                    End If
                Next intIndex
                
                If intPercentTrack = intNumPotentialElementCount Then
                    ' Matching compound, so add to list
                    RecursiveConstructAndVerifyCompound "", strPotentialElements(), dblPotentialElementStats(), intNumPotentialElementCount, intNewPotentialElementPointers(), intPointerCount + 1, dblNewMassTotal, 0, 0, sngChargeTotal, 0
                End If
                
            End If
            
            ' Update status
            lngRecursiveFunctionCallCount = lngRecursiveFunctionCallCount + 1
            If lngRecursiveFunctionCallCount Mod 500 = 0 Then
                DoEvents
                If gKeyPressAbortFormulaFinder >= 2 Or lstResults.ListCount >= CIntSafe(txtHits.Text) Then
                    Exit Sub
                End If
            End If
            If intPointerCount = 3 Then UpdateStatus intNumPotentialElementCount, intPointerCount
                                    
            ' Haven't reached dblMaxFormulaWeight
            ' Now recursively call this sub
            RecursivePCompFinder strPotentialElements(), dblPotentialElementStats(), intCurrentIndex, intNumPotentialElementCount, intNewPotentialElementPointers(), intPointerCount + 1, dblNewMassTotal, dblTargetPercents(), dblMaxFormulaWeight, sngChargeTotal
            
        End If
    Next intCurrentIndex

    Exit Sub

RecursivePCompFinderErrorHandler:
    GeneralErrorHandler "frmFinder|RecursivePCompFinder", Err.Number, Err.Description

End Sub

Public Function GetMaxElementCount() As Integer
    GetMaxElementCount = MAX_MATCHINGELEMENTS
End Function

Public Sub LoadDynamicTextCaptions()
    Dim intIndex As Integer, strCaption As String
    
    strCaption = LookupLanguageCaption(10260, "Percent")
    For intIndex = 0 To txtPercent.Count - 1
        With txtPercent(intIndex)
            ' If current value is not a number (including being a %), then
            '  replace with correct text or symbol for new langugae
            If Not IsNumeric(.Text) Then
                .Text = strCaption
            End If
        End With
    Next intIndex
    
    strCaption = LookupLanguageCaption(10270, "# or Element or Abbrev.")
    For intIndex = 0 To txtWeight.Count - 1
        With txtWeight(intIndex + 4)
            ' Look for custom weight starting with a #, or containing a space,
            '  or starting with Ý (Russian abbreviation)
            ' If found, replace with correct phrase for new language
            If Left(.Text, 1) = "#" Or InStr(.Text, " ") > 0 Or .Text = "Ý" Then
                .Text = strCaption
            End If
        End With
    Next intIndex
End Sub

Private Sub MultipleSearchMath(intNumPotentialElementCount As Integer, intMultipleSearchMin As Integer, intMultipleSearchMax As Integer)
    
    intMultipleSearchMin = CIntSafe(frmFinderOptions.txtChargeMin)
    intMultipleSearchMax = CIntSafe(frmFinderOptions.txtChargeMax)
                    
    intMultipleSearchMax = AbsMax(intMultipleSearchMin, intMultipleSearchMax)
    intMultipleSearchMin = 1
    
    If intMultipleSearchMax < intMultipleSearchMin Then intMultipleSearchMax = intMultipleSearchMin
    
    ' Initialize the status box
    UpdateStatus intNumPotentialElementCount, 3, True, intMultipleSearchMax

End Sub

Private Function ParseNum(strWork As String, ByRef intNumLength As Integer, Optional blnAllowNegative As Boolean = False) As Double
    ' Looks for a number and returns it if found
    ' If an error is found, it returns a negative number for the error code and sets intNumLength = 0
    '  -1 = No number
    '  -2 =                                             (unused)
    '  -3 = No number at all or (more likely) no number after decimal point
    '  -4 = More than one decimal point
    
    Dim strWorking As String, strFoundNum As String, intIndex As Integer, intDecPtCount As Integer

    If mDecimalSeparator = "" Then
        mDecimalSeparator = DetermineDecimalPoint()
    End If
    
    ' Set intNumLength to -1 for now
    ' If it doesn't get set to 0 (due to an error), it will get set to the
    '   length of the matched number before exiting the sub
    intNumLength = -1
    
    If strWork = "" Then strWork = EMPTY_STRINGCHAR
    If (Asc(Left(strWork, 1)) < 48 Or Asc(Left(strWork, 1)) > 57) And _
       Left(strWork, 1) <> mDecimalSeparator And _
       Not (Left(strWork, 1) = "-" And blnAllowNegative = True) Then
        intNumLength = 0          ' No number found
        ParseNum = -1
    Else
        ' Start of string is a number or a decimal point, or (if allowed) a negative sign
        For intIndex = 1 To Len(strWork)
            strWorking = Mid(strWork, intIndex, 1)
            If IsNumeric(strWorking) Or strWorking = mDecimalSeparator Or _
               (blnAllowNegative = True And strWorking = "-") Then
                strFoundNum = strFoundNum & strWorking
            Else
                Exit For
            End If
        Next intIndex
    
        If Len(strFoundNum) = 0 Or strFoundNum = mDecimalSeparator Then
            ' No number at all or (more likely) no number after decimal point
            strFoundNum = -3
            intNumLength = 0
        Else
            ' Check for more than one decimal point (. or ,)
            intDecPtCount = 0
            For intIndex = 1 To Len(strFoundNum)
                If Mid(strFoundNum, intIndex, 1) = mDecimalSeparator Then intDecPtCount = intDecPtCount + 1
            Next intIndex
            If intDecPtCount > 1 Then
                ' more than one intDecPtCount
                strFoundNum = -4
                intNumLength = 0
            Else
                ' All is fine
            End If
        End If
    
        If intNumLength < 0 Then intNumLength = Len(strFoundNum)
        ParseNum = CDblSafe(strFoundNum)
    End If
    
End Function


Private Sub PositionFormControls()
    Dim intIndex As Integer

    lstResults.Visible = True
    rtfResults.Visible = False
    
    lblPercentMaxWeight.Visible = False
    txtPercentMaxWeight.Visible = False
    lblPercentTolerance.Visible = False
    txtPercentTolerance.Visible = False
    lblMin.Visible = False
    lblMax.Visible = False
    lblPercent.Visible = False
    lblWeight.Visible = False
    
    For intIndex = 0 To 9
        txtMin(intIndex).Visible = False
        txtMax(intIndex).Visible = False
        chkElements(intIndex).Visible = True
        txtPercent(intIndex).Visible = False
        If intIndex >= 4 Then txtWeight(intIndex).Visible = False
    Next intIndex
    
End Sub

Private Sub PositionLblWeight()
    Dim intIndex As Integer
    
    ' Determine position of lblWeight
    For intIndex = 4 To 9
        If cChkBox(chkElements(intIndex)) Then
            ' Position lblWeight above this custom element
            lblWeight.Top = chkElements(intIndex).Top - chkElements(intIndex).Height - 180
            Exit For
        End If
    Next intIndex

End Sub

Private Sub OldFormulaFinder(strPotentialElements() As String, dblPotentialElementStats() As Double, intNumPotentialElementCount As Integer, dblTargetWeight As Double, dblTolerance As Double, intMultipleSearchMin As Integer, intMultipleSearchMax As Integer, intRange() As Integer, dblMaxFormulaWeight As Double, intMaxHits As Integer, dblTargetPercents() As Double)
    
    ' Iteration variables
    Dim j As Long, k As Long, l As Long, m As Long, N As Long, O As Long, P As Long, q As Long, r As Long, S As Long
    
    Dim dblTotalWeight As Double, sngTotalCharge As Single
    Dim intSubTrack As Integer, intIndex As Integer
    Dim blnHOK As Boolean
    Dim dblMultipleSearchMaxWeight As Double, dblMatchWeight As Double
    
    Dim strConstruct As String, intCurrentCharge As Integer
    
    Dim Percent(MAX_MATCHINGELEMENTS) As Double          ' The calculated percentages for the specific compound
    
On Error GoTo OldFormulaFinderErrorHandler

    dblMultipleSearchMaxWeight = dblTargetWeight * intMultipleSearchMax
    
    ' Determine the valid compounds
    For j = intRange(0, 0) To intRange(0, 1)
        For k = intRange(1, 0) To intRange(1, 1)
            For l = intRange(2, 0) To intRange(2, 1)
                For m = intRange(3, 0) To intRange(3, 1)
                    For N = intRange(4, 0) To intRange(4, 1)
                        For O = intRange(5, 0) To intRange(5, 1)
                            For P = intRange(6, 0) To intRange(6, 1)
                                For q = intRange(7, 0) To intRange(7, 1)
                                    For r = intRange(8, 0) To intRange(8, 1)
                                        For S = intRange(9, 0) To intRange(9, 1)
                                            
                                            If k Mod 10 = 0 Then
                                                DoEvents
                                            End If
    
                                            dblTotalWeight = j * dblPotentialElementStats(0, 0) + k * dblPotentialElementStats(1, 0) + l * dblPotentialElementStats(2, 0) + m * dblPotentialElementStats(3, 0) + N * dblPotentialElementStats(4, 0) + O * dblPotentialElementStats(5, 0) + P * dblPotentialElementStats(6, 0) + q * dblPotentialElementStats(7, 0) + r * dblPotentialElementStats(8, 0) + S * dblPotentialElementStats(9, 0)
                                            sngTotalCharge = j * dblPotentialElementStats(0, 1) + k * dblPotentialElementStats(1, 1) + l * dblPotentialElementStats(2, 1) + m * dblPotentialElementStats(3, 1) + N * dblPotentialElementStats(4, 1) + O * dblPotentialElementStats(5, 1) + P * dblPotentialElementStats(6, 1) + q * dblPotentialElementStats(7, 1) + r * dblPotentialElementStats(8, 1) + S * dblPotentialElementStats(9, 1)
                                            
                                            If optType(1).value = True Then
                                                ' Matching Percent Compositions
                                                If dblTotalWeight > 0 And dblTotalWeight <= dblMaxFormulaWeight Then
                                                    Percent(0) = j * dblPotentialElementStats(0, 0) / dblTotalWeight * 100
                                                    Percent(1) = k * dblPotentialElementStats(1, 0) / dblTotalWeight * 100
                                                    Percent(2) = l * dblPotentialElementStats(2, 0) / dblTotalWeight * 100
                                                    Percent(3) = m * dblPotentialElementStats(3, 0) / dblTotalWeight * 100
                                                    Percent(4) = N * dblPotentialElementStats(4, 0) / dblTotalWeight * 100
                                                    Percent(5) = O * dblPotentialElementStats(5, 0) / dblTotalWeight * 100
                                                    Percent(6) = P * dblPotentialElementStats(6, 0) / dblTotalWeight * 100
                                                    Percent(7) = q * dblPotentialElementStats(7, 0) / dblTotalWeight * 100
                                                    Percent(8) = r * dblPotentialElementStats(8, 0) / dblTotalWeight * 100
                                                    Percent(9) = S * dblPotentialElementStats(9, 0) / dblTotalWeight * 100
                                                            
                                                    intSubTrack = 0
                                                    For intIndex = 0 To intNumPotentialElementCount - 1
                                                        If Percent(intIndex) >= dblTargetPercents(intIndex, 0) And Percent(intIndex) <= dblTargetPercents(intIndex, 1) Then intSubTrack = intSubTrack + 1
                                                    Next intIndex
                                                    If intSubTrack = intNumPotentialElementCount Then
                                                        ' All of the elements have percent compositions matching the target
                                                        
                                                        ' Call subroutine to construct formula and verify hydrogens
                                                        strConstruct = ""      ' strConstruct is passed by reference
                                                        blnHOK = ConstructAndVerifyCompound(strConstruct, j, k, l, m, N, O, P, q, r, S, strPotentialElements(0), strPotentialElements(1), strPotentialElements(2), strPotentialElements(3), strPotentialElements(4), strPotentialElements(5), strPotentialElements(6), strPotentialElements(7), strPotentialElements(8), strPotentialElements(9), dblTotalWeight, dblTargetWeight, dblTolerance, sngTotalCharge, 0)
                                                        ' No need to set blnHOK to True when not verifying hydrogens since program does this in ConstructAndVerifyCompound function
                                                        
                                                        If strConstruct <> "" And blnHOK Then
                                                            ConstructResultsLine strConstruct, dblTotalWeight, -1, sngTotalCharge
    
                                                            ' Add % Comp results to list
                                                            strConstruct = strConstruct & " " & LookupLanguageCaption(10535, "has") & " "
                                                            If j <> 0 Then strConstruct = strConstruct & strPotentialElements(0) & "=" & Format(Percent(0), "##.##") & "% "
                                                            If k <> 0 Then strConstruct = strConstruct & strPotentialElements(1) & "=" & Format(Percent(1), "##.##") & "% "
                                                            If l <> 0 Then strConstruct = strConstruct & strPotentialElements(2) & "=" & Format(Percent(2), "##.##") & "% "
                                                            If m <> 0 Then strConstruct = strConstruct & strPotentialElements(3) & "=" & Format(Percent(3), "##.##") & "% "
                                                            If N <> 0 Then strConstruct = strConstruct & strPotentialElements(4) & "=" & Format(Percent(4), "##.##") & "% "
                                                            If O <> 0 Then strConstruct = strConstruct & strPotentialElements(5) & "=" & Format(Percent(5), "##.##") & "% "
                                                            If P <> 0 Then strConstruct = strConstruct & strPotentialElements(6) & "=" & Format(Percent(6), "##.##") & "% "
                                                            If q <> 0 Then strConstruct = strConstruct & strPotentialElements(7) & "=" & Format(Percent(7), "##.##") & "% "
                                                            If r <> 0 Then strConstruct = strConstruct & strPotentialElements(8) & "=" & Format(Percent(8), "##.##") & "% "
                                                            If S <> 0 Then strConstruct = strConstruct & strPotentialElements(9) & "=" & Format(Percent(9), "##.##") & "% "
                                                            lstResults.AddItem strConstruct
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                ' Matching Molecular Weights

                                                If dblTotalWeight <= dblMultipleSearchMaxWeight + dblTolerance Then
                                                    For intCurrentCharge = intMultipleSearchMin To intMultipleSearchMax
                                                        dblMatchWeight = dblTargetWeight * intCurrentCharge
                                                        If dblTotalWeight <= dblMatchWeight + dblTolerance And _
                                                           dblTotalWeight >= dblMatchWeight - dblTolerance Then
                                                            ' Within dblTolerance
                                                            ' Call subroutine to construct formula and verify hydrogens
                                                            strConstruct = ""      ' strConstruct is passed by reference
                                                            blnHOK = ConstructAndVerifyCompound(strConstruct, j, k, l, m, N, O, P, q, r, S, strPotentialElements(0), strPotentialElements(1), strPotentialElements(2), strPotentialElements(3), strPotentialElements(4), strPotentialElements(5), strPotentialElements(6), strPotentialElements(7), strPotentialElements(8), strPotentialElements(9), dblTotalWeight, dblTargetWeight * intCurrentCharge, dblTolerance, sngTotalCharge, intCurrentCharge)
                                                            ' No need to set blnHOK = 1 when not verifying hydrogens since program does this in ConstructAndVerifyCompound function
                                        
                                                            If strConstruct <> "" And blnHOK Then
                                                                ConstructResultsLine strConstruct, dblTotalWeight, dblTargetWeight, sngTotalCharge
                                        
                                                                lstResults.AddItem strConstruct
                                                            End If
                                                           Exit For
                                                        End If
                                                    Next intCurrentCharge
                                                Else

                                                    ' Jump out of loop since weight is too high
                                                    ' Determine which variable is causing the weight to be too high
                                                    ' Incrementing "s" would definitely make the weight too high, so set it to its max (so it will zero and increment "r")
                                                    S = intRange(9, 1)
                                                    If (j * dblPotentialElementStats(0, 0) + k * dblPotentialElementStats(1, 0) + l * dblPotentialElementStats(2, 0) + m * dblPotentialElementStats(3, 0) + N * dblPotentialElementStats(4, 0) + O * dblPotentialElementStats(5, 0) + P * dblPotentialElementStats(6, 0) + q * dblPotentialElementStats(7, 0) + (r + 1) * dblPotentialElementStats(8, 0)) > (dblTolerance + dblMultipleSearchMaxWeight) Then
                                                        ' Incrementing r would make the weight too high, so set it to its max (so it will zero and increment q)
                                                        r = intRange(8, 1)
                                                        If (j * dblPotentialElementStats(0, 0) + k * dblPotentialElementStats(1, 0) + l * dblPotentialElementStats(2, 0) + m * dblPotentialElementStats(3, 0) + N * dblPotentialElementStats(4, 0) + O * dblPotentialElementStats(5, 0) + P * dblPotentialElementStats(6, 0) + (q + 1) * dblPotentialElementStats(7, 0)) > (dblTolerance + dblMultipleSearchMaxWeight) Then
                                                            q = intRange(7, 1)
                                                            If (j * dblPotentialElementStats(0, 0) + k * dblPotentialElementStats(1, 0) + l * dblPotentialElementStats(2, 0) + m * dblPotentialElementStats(3, 0) + N * dblPotentialElementStats(4, 0) + O * dblPotentialElementStats(5, 0) + (P + 1) * dblPotentialElementStats(6, 0)) > (dblTolerance + dblMultipleSearchMaxWeight) Then
                                                                P = intRange(6, 1)
                                                                If (j * dblPotentialElementStats(0, 0) + k * dblPotentialElementStats(1, 0) + l * dblPotentialElementStats(2, 0) + m * dblPotentialElementStats(3, 0) + N * dblPotentialElementStats(4, 0) + (O + 1) * dblPotentialElementStats(5, 0)) > (dblTolerance + dblMultipleSearchMaxWeight) Then
                                                                    O = intRange(5, 1)
                                                                    If (j * dblPotentialElementStats(0, 0) + k * dblPotentialElementStats(1, 0) + l * dblPotentialElementStats(2, 0) + m * dblPotentialElementStats(3, 0) + (N + 1) * dblPotentialElementStats(4, 0)) > (dblTolerance + dblMultipleSearchMaxWeight) Then
                                                                        N = intRange(4, 1)
                                                                        If (j * dblPotentialElementStats(0, 0) + k * dblPotentialElementStats(1, 0) + l * dblPotentialElementStats(2, 0) + (m + 1) * dblPotentialElementStats(3, 0)) > (dblTolerance + dblMultipleSearchMaxWeight) Then
                                                                            m = intRange(3, 1)
                                                                            If (j * dblPotentialElementStats(0, 0) + k * dblPotentialElementStats(1, 0) + (l + 1) * dblPotentialElementStats(2, 0)) > (dblTolerance + dblMultipleSearchMaxWeight) Then
                                                                                l = intRange(2, 1)
                                                                                If (j * dblPotentialElementStats(0, 0) + (k + 1) * dblPotentialElementStats(1, 0)) > (dblTolerance + dblMultipleSearchMaxWeight) Then
                                                                                    k = intRange(1, 1)
                                                                                    If ((j + 1) * dblPotentialElementStats(0, 0)) > (dblTolerance + dblMultipleSearchMaxWeight) Then
                                                                                        j = intRange(0, 1)
                                                                                    End If
                                                                                End If
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            End If
                                            If lstResults.ListCount >= intMaxHits Or gKeyPressAbortFormulaFinder >= 2 Then
                                                ' Set variables to their maximum so all the loops will end
                                                j = intRange(0, 1)
                                                k = intRange(1, 1)
                                                l = intRange(2, 1)
                                                m = intRange(3, 1)
                                                N = intRange(4, 1)
                                                O = intRange(5, 1)
                                                P = intRange(6, 1)
                                                q = intRange(7, 1)
                                                r = intRange(8, 1)
                                                S = intRange(9, 1)
                                            End If
                                            If gKeyPressAbortFormulaFinder >= 2 Then Exit Sub
                                        Next S
                                    Next r
                                Next q
                            Next P
                        Next O
                    Next N
                Next m
            Next l
        Next k
        If intRange(0, 1) <> 0 Then
            If intMultipleSearchMin = 0 Then
                lblPercentCompleted = LookupLanguageCaption(10120, "Searching") & " " & _
                                      Str(CIntSafeDbl(j / intRange(0, 1) * 100)) & _
                                      LookupLanguageCaption(10110, "% Completed")
            Else
                lblPercentCompleted = LookupLanguageCaption(10120, "Searching") & " " & _
                                      Str(CIntSafeDbl(j / (intRange(0, 1) * intMultipleSearchMax) * 100)) & _
                                      LookupLanguageCaption(10110, "% Completed")
            End If
        End If
    
    Next j

    Exit Sub

OldFormulaFinderErrorHandler:
    GeneralErrorHandler "frmFinder|OldFormulaFinder", Err.Number, Err.Description

End Sub

Private Sub RecursiveConstructAndVerifyCompound(strResults As String, strPotentialElements() As String, dblPotentialElementStats() As Double, intNumPotentialElementCount As Integer, intPotentialElementPointers() As Integer, intPointerCount As Integer, dblTotalMass As Double, dblTargetMass As Double, dblWeightTolerance As Double, ByVal sngTotalCharge As Single, intMultipleMtoZCharge As Integer)
    ' common function to both molecular weight and percent composition matching
    ' strResults is blank for molecular weight matching, but has the % compositions for %comp matching
    ' if intMultipleMtoZCharge is 0 then searching for target weights, otherwise searching for target m/z's
    
    Dim intIndex As Integer, maxH%, Y%, swapVal%
    Dim blnHyrdrogenOK As Boolean, blnChargeOK As Boolean
    
    Dim intElementCount() As Integer
    Dim strEmpiricalResult() As String
    Dim intEmpiricalResultElementCount() As Integer
    Dim strSwap As String
    
    ReDim intElementCount(intNumPotentialElementCount)    ' Store the Number of each element
    ReDim strEmpiricalResult(intNumPotentialElementCount)               ' Store the elements and abbreviations of the found formula so that they can be properly sequenced according to empirical formula conventions
    ReDim intEmpiricalResultElementCount(intNumPotentialElementCount)   ' Store the Number of each element and abbreviation for empirical formula conversion
    
    ' Counters for elements of interest (hydrogen, carbon, silicon, nitrogen, phosphorus, chlorine, iodine, flourine, bromine, and other)
    Dim udtElementNum As udtElementNumType
    
On Error GoTo RecursiveConstructAndVerifyCompoundErrorHandler

    For intIndex = 0 To intPointerCount - 1
        ' For each pointer, increment intElementCount
        intElementCount(intPotentialElementPointers(intIndex)) = intElementCount(intPotentialElementPointers(intIndex)) + 1
    Next intIndex
    
    If Not cChkBox(frmFinderOptions.chkSort) Then
        ' Don't sort or convert to formatted empirical formula
        For intIndex = 0 To intNumPotentialElementCount - 1
            ' Step through elements, if intElementCount > 0, then add to formula
            If intElementCount(intIndex) <> 0 Then strResults = strResults & strPotentialElements(intIndex)
            If intElementCount(intIndex) > 1 Then strResults = strResults & LTrim(CStr(intElementCount(intIndex)))
        Next intIndex
    Else
        
        ' Convert to empirical formula and sort
        For intIndex = 0 To intNumPotentialElementCount - 1
            If intElementCount(intIndex) <> 0 Then strEmpiricalResult(intIndex) = strPotentialElements(intIndex)
            If intElementCount(intIndex) >= 1 Then intEmpiricalResultElementCount(intIndex) = intElementCount(intIndex)
        Next intIndex
        
        ' First find C
        For intIndex = 0 To intNumPotentialElementCount - 1
            If strEmpiricalResult(intIndex) = "C" Then
                strResults = strResults & strEmpiricalResult(intIndex)       ' i.e. C
                If intEmpiricalResultElementCount(intIndex) > 1 Then strResults = strResults & _
                                                         LTrim(CStr(intEmpiricalResultElementCount(intIndex)))
                ' blank this strEmpiricalResult so it doesn't get repeated
                strEmpiricalResult(intIndex) = ""
                Exit For
            End If
        Next intIndex
        For intIndex = 0 To intNumPotentialElementCount - 1
            If strEmpiricalResult(intIndex) = "H" Then
                strResults = strResults & strEmpiricalResult(intIndex)       ' i.e. H
                If intEmpiricalResultElementCount(intIndex) > 1 Then strResults = strResults & _
                                                         LTrim(CStr(intEmpiricalResultElementCount(intIndex)))
                ' blank this strEmpiricalResult so it doesn't get repeated
                strEmpiricalResult(intIndex) = ""
                Exit For
            End If
        Next intIndex
        
        ' Alphabatize the remaining elements/abbreviations via a simple bubble sort
        For Y = intNumPotentialElementCount - 1 To 1 Step -1       ' Sort from end to start
            For intIndex = 0 To Y - 1
                If strEmpiricalResult(intIndex) > strEmpiricalResult(intIndex + 1) Then
                    ' strSwap the elements/abbreviations
                    strSwap = strEmpiricalResult(intIndex)
                    strEmpiricalResult(intIndex) = strEmpiricalResult(intIndex + 1)
                    strEmpiricalResult(intIndex + 1) = strSwap
                    ' and their intEmpiricalResultElementCount values
                    swapVal = intEmpiricalResultElementCount(intIndex)
                    intEmpiricalResultElementCount(intIndex) = intEmpiricalResultElementCount(intIndex + 1)
                    intEmpiricalResultElementCount(intIndex + 1) = swapVal
                End If
            Next intIndex
        Next Y
        
        ' Now insert the alphabatized elements into the string
        For intIndex = 0 To intNumPotentialElementCount - 1
            ' If C or H was found above, its strEmpiricalResult was made blank, so need the following
            ' check to exclude C and H
            If strEmpiricalResult(intIndex) <> "" Then
                strResults = strResults & strEmpiricalResult(intIndex)
                If intEmpiricalResultElementCount(intIndex) > 1 Then strResults = strResults & _
                                                         LTrim(CStr(intEmpiricalResultElementCount(intIndex)))
            End If
        Next intIndex

    End If
    
    ' Verify hydrogens if requested
    If cChkBox(frmFinderOptions.chkVerifyHydrogens) Or intMultipleMtoZCharge > 0 Then
        
        ' Determine Number of C, Si, N, P, O, S, Cl, I, F, Br and H atoms
        For intIndex = 0 To intNumPotentialElementCount - 1
            Select Case strPotentialElements(intIndex)
            Case "C": udtElementNum.C = udtElementNum.C + intElementCount(intIndex)
            Case "Si": udtElementNum.Si = udtElementNum.Si + intElementCount(intIndex)
            Case "N": udtElementNum.N = udtElementNum.N + intElementCount(intIndex)
            Case "P": udtElementNum.P = udtElementNum.P + intElementCount(intIndex)
            Case "O": udtElementNum.O = udtElementNum.O + intElementCount(intIndex)
            Case "S": udtElementNum.S = udtElementNum.S + intElementCount(intIndex)
            Case "Cl": udtElementNum.Cl = udtElementNum.Cl + intElementCount(intIndex)
            Case "I": udtElementNum.I = udtElementNum.I + intElementCount(intIndex)
            Case "F": udtElementNum.F = udtElementNum.F + intElementCount(intIndex)
            Case "Br": udtElementNum.Br = udtElementNum.Br + intElementCount(intIndex)
            Case "H": udtElementNum.H = udtElementNum.H + intElementCount(intIndex)
            Case Else: udtElementNum.Other = udtElementNum.Other + intElementCount(intIndex)
            End Select
        Next intIndex
        
        ' Compute maximum Number of hydrogens
        If udtElementNum.Si = 0 And udtElementNum.C = 0 And udtElementNum.N = 0 And _
           udtElementNum.P = 0 And udtElementNum.Other = 0 And (udtElementNum.O > 0 Or udtElementNum.S > 0) Then
            ' Only O and S
            maxH = 3
        Else
            ' Formula is: [#C*2 + 3 - (2 if N or P present)] + [#N + 3 - (1 if C or Si present)] + [#other elements * 4 + 3], where we assume other elements can have a coordination Number of up to 7
            If udtElementNum.C > 0 Or udtElementNum.Si > 0 Then
                maxH = maxH + (udtElementNum.C + udtElementNum.Si) * 2 + 3
'                If udtElementNum.N > 0 Or udtElementNum.P > 0 Then maxh = maxh - 2
            End If
            If udtElementNum.N > 0 Or udtElementNum.P > 0 Then
                maxH = maxH + (udtElementNum.N + udtElementNum.P) + 3
'                If udtElementNum.C > 0 Or udtElementNum.Si > 0 Then maxh = maxh - 1
            End If
            ' Correction for carbon contribution
            ' Combine the above two remarked if's into:
            If (udtElementNum.N > 0 Or udtElementNum.P > 0) And (udtElementNum.C > 0 Or udtElementNum.Si > 0) Then maxH = maxH - 3
            
            If udtElementNum.Other > 0 Then maxH = maxH + udtElementNum.Other * 4 + 3
        End If
        
        ' correct for if H only
        If maxH < 3 Then maxH = 3
        
        ' correct for halogens
        maxH = maxH - udtElementNum.F - udtElementNum.Cl - udtElementNum.Br - udtElementNum.I
        
        ' correct for negative maxh
        If maxH < 0 Then maxH = 0
        
        ' Verify H's
        If udtElementNum.H <= maxH Or Not cChkBox(frmFinderOptions.chkVerifyHydrogens) Then
            blnHyrdrogenOK = True
        Else
            blnHyrdrogenOK = False
        End If
    Else
        blnHyrdrogenOK = True
    End If
    
    ' Only proceed if compound isn't blank and hydrogens check out
    If strResults <> "" And blnHyrdrogenOK Then
        
        ' See if sngTotalCharge is within charge limits
        sngTotalCharge = CorrectChargeEmpirical(sngTotalCharge, udtElementNum, blnChargeOK)
        
        ' If charge is within range and checking for multiples, see if correct m/z too
        If blnChargeOK And intMultipleMtoZCharge > 0 Then
            CheckMtoZWithTarget dblTotalMass, sngTotalCharge, dblTargetMass, dblWeightTolerance, intMultipleMtoZCharge, blnChargeOK
        End If
    
        ' Only proceed if charge checks out
        If blnChargeOK Then
            If optType(0).value = True Then
                ' Matching Molecular Weights
                ConstructResultsLine strResults, dblTotalMass, dblTargetMass, sngTotalCharge
            Else
                ' Matching Percent Compositions
                ConstructResultsLine strResults, dblTotalMass, -1, sngTotalCharge
                
                strResults = strResults & " " & LookupLanguageCaption(10535, "has") & " "
                ' Add % info
                For intIndex = 0 To intNumPotentialElementCount - 1
                    If intElementCount(intIndex) <> 0 Then
                        strResults = strResults & strPotentialElements(intIndex) & "=" & _
                                     Format$(intElementCount(intIndex) * dblPotentialElementStats(intIndex, 0) / dblTotalMass * 100, "##.##") & "% "
                    End If
                Next intIndex
            
            End If
            lstResults.AddItem strResults
        End If
    End If
    
    Exit Sub

RecursiveConstructAndVerifyCompoundErrorHandler:
    GeneralErrorHandler "frmFinder|RecursiveConstructAndVerifyCompound", Err.Number, Err.Description

End Sub

Public Sub ResizeForm(Optional blnEnlargeToMinimums As Boolean = False)
    Dim X%, lngPreferredWidth%
    Dim lngMinHeight As Long
    Dim lngTopPosAdjust As Long

    Dim lngEffectiveFormWidth As Long
    Dim lngEffectiveFormHeight As Long
    
    lngMinHeight = 6725

On Error GoTo ResizeFormErrorHandler

    With frmFinder
        If .WindowState <> vbMinimized Then
            ' define minimum width
            lngPreferredWidth = 1.75 * TextWidth(rtfResults.Text)
            
            If lngPreferredWidth < 9100 Then lngPreferredWidth = 9100
            If lngPreferredWidth > Screen.Width Then lngPreferredWidth = Screen.Width
            
            If .Width < lngPreferredWidth Then
                If blnEnlargeToMinimums And .WindowState = vbNormal Then
                    .Width = lngPreferredWidth
                End If
                lngEffectiveFormWidth = lngPreferredWidth
            Else
                lngEffectiveFormWidth = .Width
            End If
            
            If .Height < lngMinHeight Then
                If blnEnlargeToMinimums And .WindowState = vbNormal Then
                    .Height = lngMinHeight
                End If
                lngEffectiveFormHeight = lngMinHeight
            Else
                lngEffectiveFormHeight = .Height
            End If
        End If
    End With

    If frmFinder.WindowState <> vbMinimized Then
        lngTopPosAdjust = lngEffectiveFormHeight - lngMinHeight
        
        lblInstruct.Top = 120
        lblInstruct.Left = 120
        
        With lstResults
            .Top = 1200
            .Left = 120
            .Height = 3400 + lngTopPosAdjust
            
            .Width = lngEffectiveFormWidth - 5400
            If frmFinderOptions.cboSearchType.ListIndex = 0 Then .Width = .Width + 1250
            
            .Font = objMwtWin.RtfFontName
            .FontSize = objMwtWin.RtfFontSize
        End With
        
        With rtfResults
            .Top = lstResults.Top
            .Left = 120
            .Height = 3200 + lngTopPosAdjust
            
            .Width = lstResults.Width
            .Font.Name = objMwtWin.RtfFontName
            .Font.Size = objMwtWin.RtfFontSize
        End With
        
        lblMWT.Top = 4625 + lngTopPosAdjust
        lblMWT.Left = 120
        txtMWT.Top = lblMWT.Top
        txtMWT.Left = 2900
        lblWtTolerance.Top = 5075 + lngTopPosAdjust
        lblWtTolerance.Left = lblMWT.Left
        txtWeightTolerance.Top = lblWtTolerance.Top
        txtWeightTolerance.Left = txtMWT.Left
        
        lblPercentMaxWeight.Top = lblMWT.Top
        lblPercentMaxWeight.Left = lblMWT.Left
        txtPercentMaxWeight.Top = txtMWT.Top
        txtPercentMaxWeight.Left = txtMWT.Left
        lblPercentTolerance.Top = lblWtTolerance.Top
        lblPercentTolerance.Left = lblWtTolerance.Left
        txtPercentTolerance.Top = txtWeightTolerance.Top
        txtPercentTolerance.Left = txtWeightTolerance.Left
        txtPercentTolerance.Width = 735
        
        chkPpmMode.Top = lblWtTolerance.Top
        chkPpmMode.Left = 4200
        
        chkShowDeltaMass.Top = chkPpmMode.Top + 255
        chkShowDeltaMass.Left = chkPpmMode.Left
        
        optType(0).Top = 5475 + lngTopPosAdjust
        optType(1).Top = 5825 + lngTopPosAdjust
        
        cmdCalculate.Top = 5625 + lngTopPosAdjust
        cmdAbort.Top = cmdCalculate.Top
        cmdPrint.Top = cmdCalculate.Top
        cmdCopy.Top = cmdCalculate.Top
        cmdCopyAsRTF.Top = lblWtTolerance.Top - 50
        cmdDisplayIsotopicAbundance.Top = cmdCopyAsRTF.Top
        cmdOK.Top = cmdCalculate.Top
            
        cmdCalculate.Left = lngEffectiveFormWidth - 5260
        cmdAbort.Left = cmdCalculate.Left
        cmdPrint.Left = lngEffectiveFormWidth - 3940
        cmdCopy.Left = lngEffectiveFormWidth - 2620
        cmdCopyAsRTF.Left = cmdCopy.Left - (cmdCopyAsRTF.Width - cmdCopy.Width) / 2
        cmdOK.Left = lngEffectiveFormWidth - 1300
        cmdDisplayIsotopicAbundance.Left = cmdOK.Left - (cmdDisplayIsotopicAbundance.Width - cmdOK.Width) / 2
        
        lblElement.Top = 720
        lblElement.Width = 1215
        lblElement.Left = lngEffectiveFormWidth - 3650
        lblPercentCompleted.Left = lngEffectiveFormWidth - 1350
        lblPercentCompleted.Top = 1500
        
        ' Hidden Labels
        lblMin.Top = lblElement.Top
        lblMax.Top = lblElement.Top
        lblPercent.Top = lblElement.Top
        lblMin.Left = lngEffectiveFormWidth - 5150
        lblMax.Left = lngEffectiveFormWidth - 4450
        lblPercent.Left = lngEffectiveFormWidth - 2100
        lblPercent.Width = 615
        Call PositionLblWeight
        lblWeight.Left = lblPercentCompleted.Left + 120
        lblWeight.Width = 855
    
        For X = 0 To 9
           txtMin(X).Top = X * 360 + lstResults.Top
           txtMin(X).Left = lblMin.Left
           txtMax(X).Top = X * 360 + lstResults.Top
           txtMax(X).Left = lblMax.Left
           chkElements(X).Top = X * 360 + lstResults.Top
           chkElements(X).Left = lngEffectiveFormWidth - 3800
           chkElements(X).Width = 1700
           txtPercent(X).Top = X * 360 + lstResults.Top
           txtPercent(X).Left = lblPercent.Left
           txtPercent(X).Width = 735
        Next X
    
        chkElements(0).Caption = LookupLanguageCaption(10400, "Carbon")
        chkElements(1).Caption = LookupLanguageCaption(10405, "Hydrogen")
        chkElements(2).Caption = LookupLanguageCaption(10410, "Nitrogen")
        chkElements(3).Caption = LookupLanguageCaption(10415, "Oxygen")
        For X = 4 To 9
            chkElements(X).Caption = LookupLanguageCaption(10420, "Custom") & Trim(Str(X - 3)) & "_"
        Next X
        
        txtHits.Top = txtPercent(0).Top
        txtHits.Left = lblPercentCompleted.Left
        lblHits.Top = lblPercent.Top
        lblHits.Left = lblPercentCompleted.Left
        
        cmdFinderOptions.Top = 50
        cmdFinderOptions.Left = lngEffectiveFormWidth - 1600
        
        For X = 4 To 9
            txtWeight(X).Top = X * 360 + lstResults.Top
            txtWeight(X).Left = lblPercentCompleted.Left
        Next X
    
    End If
    
    Exit Sub

ResizeFormErrorHandler:
    Debug.Assert False
    GeneralErrorHandler "frmFinder|ResizeForm", Err.Number, Err.Description

End Sub

Private Function ResultsToRtf(strWorkText As String) As String
    Dim strWorkChar As String, strRTF As String
    Dim X As Integer, Y As Integer
    Dim strChargeString As String
    
    ' Converts plain text to formatted rtf text.
    ' Rtf string must begin with {{\fonttbl{\f0\fcharset0\fprq2 Times New Roman;}}\pard\plain\fs25
    ' and must end with }
    
    strRTF = ""
    
    For X = 1 To Len(strWorkText)
        strWorkChar = Mid(strWorkText, X, 1)
        If IsNumeric(strWorkChar) Or strWorkChar = glbDecimalSeparator Then
            ' Number or period, so subscript it if it's not followed by an underscore (underscore is used for custom elements: C1_, C2_, etc.)
            If Mid(strWorkText, X + 1, 1) = "_" Then
                strRTF = strRTF & strWorkChar & "_"
                X = X + 1
            Else
                If X = 1 Then
                    ' at beginning of line, so leave it alone. Probably out of place
                    strRTF = strRTF & strWorkChar
                Else
                    ' find entire number
                    Y = 1
                    Do While IsNumeric(Mid(strWorkText, X + Y, 1))
                        Y = Y + 1
                    Loop
                    strRTF = strRTF & "{\sub " & Mid(strWorkText, X, Y) & "}"
                    X = X + Y - 1
                End If
            End If
        ElseIf strWorkChar = "^" Then
            ' Found a caret, ignore it and superscript the number after it
            ' Find entire number
            Y = 1
            Do While IsNumeric(Mid(strWorkText, X + Y, 1)) Or Mid(strWorkText, X + Y, 1) = "-"
                Y = Y + 1
            Loop
            strChargeString = ChargeValueToString(CDblSafe(Mid(strWorkText, X + 1, Y - 1)))
            
            strRTF = strRTF & "{\super " & strChargeString & "}"
            X = X + Y
        ElseIf strWorkChar = "+" Then
            ' Found a plus sign, superscript it
            strRTF = strRTF & "{\super +}"
        ElseIf strWorkChar = "_" Then
            ' Found an underscore following a letter, superscript it
            strRTF = strRTF & "{\super _}"
        Else
            strRTF = strRTF & strWorkChar
        End If
    Next X

    ResultsToRtf = strRTF

End Function

Public Sub SetWeightMatchingMode(intMatchingMode As Integer)

    ' Match molecular weight has intMatchingMode = 0
    ' Match percent compositions has intMatchingMode = 1
        
    Dim intIndex As Integer
        
    If gKeyPressAbortFormulaFinder = 1 Then
        gKeyPressAbortFormulaFinder = 3
    End If
    
    If intMatchingMode = 0 Then
        ' Match molecular weights
        lblPercent.Visible = False
        lblPercentTolerance.Visible = False
        txtPercentTolerance.Visible = False
        lblPercentMaxWeight.Visible = False
        txtPercentMaxWeight.Visible = False
        lblWtTolerance.Visible = True
        txtWeightTolerance.Visible = True
        txtMWT.Visible = True
        txtMWT.Visible = True
        chkPpmMode.Visible = True
    Else
        ' Match percent compositions
        lblWtTolerance.Visible = False
        txtWeightTolerance.Visible = False
        txtMWT.Visible = False
        txtMWT.Visible = False
        chkPpmMode.Visible = False
        lblPercentMaxWeight.Visible = True
        txtPercentMaxWeight.Visible = True
        lblPercentTolerance.Visible = True
        txtPercentTolerance.Visible = True
    End If

    For intIndex = 0 To 9
        If intMatchingMode = 0 Then
            txtPercent(intIndex).Visible = False
        Else
            If cChkBox(chkElements(intIndex)) Then
                txtPercent(intIndex).Visible = True
                lblPercent.Visible = True
            End If
        End If
    Next intIndex

    frmFinderOptions.UpdateCheckBoxes

End Sub

Private Sub ShellSortResults(strCodeString() As String, udtResultStats() As udtResultStatsType, intPointerArray() As Integer, intLowIndex As Integer, intHiIndex As Integer)
    Dim eSortMode As smcFinderResultsSortModeConstants
    
    ' Variables common to both
    Dim intSwapCode As Integer, blnSwapThem As Boolean
    
    ' Sort the list via a shell sort
    Dim intMaxRow As Integer, intOffset As Integer, intLimit As Integer, intSwitch As Integer
    Dim intRow As Integer
    Dim intRemainder As Integer, intSortTrack As Integer, intSortTrackMax As Integer

    ' Do some math to help predict percent of sort completed over time
    intSortTrackMax = 0
    intRemainder = intHiIndex + 1
    intRemainder = intRemainder \ 2
    Do While intRemainder > 0
        intSortTrackMax = intSortTrackMax + 1
        intRemainder = intRemainder \ 2
    Loop
    
    eSortMode = frmFinderOptions.LookupResultsSortModeTypeByIndex(frmFinderOptions.cboSortResults.ListIndex)
        
    ' Set comparison intOffset to half the number of records
    intMaxRow = intHiIndex + 1   ' Add 1 since list goes from 0 to intHiIndex rather than 1 to intHiIndex + 1
    intOffset = intMaxRow \ 2

    Do While intOffset > 0          ' Loop until intOffset gets to zero.
        ' Check for KeyPress and keep lstResults box updated (which indicates sort progessing)
        DoEvents

        If gKeyPressAbortFormulaFinder >= 2 Then
            ' A key was pressed, so stop sorting
            If gKeyPressAbortFormulaFinder = 2 Then MsgBox LookupLanguageCaption(10450, "A key was pressed.") & "  " & LookupLanguageCaption(10470, "Stopping sort."), _
                                        vbOKOnly + vbCritical, LookupLanguageCaption(10455, "Abort")
            If gKeyPressAbortFormulaFinder = 3 Then MsgBox LookupLanguageCaption(10460, "The mouse was clicked outside the results box.") & "  " & LookupLanguageCaption(10470, "Stopping sort."), _
                                        vbOKOnly + vbCritical, LookupLanguageCaption(10455, "Abort")
            gKeyPressAbortFormulaFinder = 4
            Exit Do
        End If

        intLimit = intMaxRow - intOffset
        Do
            intSwitch = 0         ' Assume no switches at this intOffset.

            ' Compare elements and intSwitch ones out of order:
            For intRow = intLowIndex To intLimit - 1
                blnSwapThem = False
                If eSortMode = smcSortByFormula Then
                    ' Sort by Formula
                    If strCodeString(intPointerArray(intRow)) > _
                       strCodeString(intPointerArray(intRow + intOffset)) Then blnSwapThem = True
                Else
                    ' Sort numerically
                    If eSortMode = smcSortByMWT Then
                        If udtResultStats(intPointerArray(intRow)).Mass > udtResultStats(intPointerArray(intRow + intOffset)).Mass Then
                            blnSwapThem = True
                        End If
                    ElseIf eSortMode = smcSortByDeltaMass Then
                        If udtResultStats(intPointerArray(intRow)).DeltaMass > udtResultStats(intPointerArray(intRow + intOffset)).DeltaMass Then
                            blnSwapThem = True
                        End If
                    ElseIf eSortMode = smcSortByCharge Then
                        If udtResultStats(intPointerArray(intRow)).Charge > udtResultStats(intPointerArray(intRow + intOffset)).Charge Then
                            blnSwapThem = True
                        End If
                    Else
                        ' Assume eSortMode = smcSortByMZ
                        If udtResultStats(intPointerArray(intRow)).MZ > udtResultStats(intPointerArray(intRow + intOffset)).MZ Then
                            blnSwapThem = True
                        End If
                    End If
                End If
                
                If blnSwapThem = True Then
                    ' If the Low item should follow the High item, then swap them
                    ' Swap the two elements
                    intSwapCode = intPointerArray(intRow)
                    intPointerArray(intRow) = intPointerArray(intRow + intOffset)
                    intPointerArray(intRow + intOffset) = intSwapCode
                    intSwitch = intRow
                End If
            
            Next intRow

            ' Sort on next pass only to where last intSwitch was made:
            intLimit = intSwitch - intOffset
        Loop While intSwitch

        ' No switches at last intOffset, try one half as big:
        intOffset = intOffset \ 2
        
        intSortTrack = intSortTrack + 1
        If intSortTrack > 1 And intSortTrackMax <> 0 Then
            lblPercentCompleted.Caption = LookupLanguageCaption(10115, "Sorting") & " " & _
                                          Str(Format(intSortTrack / intSortTrackMax * 100, "##")) & _
                                          LookupLanguageCaption(10110, "% Completed")
            
            ' Change view of lstresults box so first line is on top, avoids screen update problem when sorting
            If lstResults.ListCount > 0 Then
                lstResults.ListIndex = 0
                lstResults.ListIndex = -1
            End If
        End If
    Loop
    
End Sub

Private Sub UpdateMinMax(Optional intUpdateLine As Integer = -1)
    Dim intIndex As Integer, intIndexStart As Integer, intIndexEnd As Integer, intMultiplier As Integer
    Dim dblElementMass As Double
    
    If cChkBox(frmFinderOptions.chkAutoSetBounds) Then
        ' Only update Min and Max if user wishes.
        If intUpdateLine = -1 Then
            intIndexStart = 0
            intIndexEnd = 9
        Else
            intIndexStart = intUpdateLine
            intIndexEnd = intUpdateLine
        End If
            
        If CDblSafe(txtMWT.Text) > 0 Then
            For intIndex = intIndexStart To intIndexEnd
                
                ' Find the mass for each element
                ' Set the maximum search value for each element to a reasonable value,
                '  based on it's mass, the search mass, and the search tolerance
                ' For example, if the search mass is 200 and the tolerance is 0.5,
                '  then for Nitrogen (mass=14), the max atom count is set to 15 since
                '  15 * 14= 210 (a little more than 200)
                If intIndex = 1 And cChkBox(frmFinderOptions.chkVerifyHydrogens) Then
                    dblElementMass = FinderMatchStats(1) * 2
                Else
                    dblElementMass = FinderMatchStats(intIndex)
                End If
                If dblElementMass <= 0 Then
                    txtMax(intIndex).Text = "20"
                Else
                    If optType(0).value = True Then
                        If Not cChkBox(frmFinderOptions.chkFindTargetMtoZ) Then
                            txtMax(intIndex).Text = Trim(Str(CIntSafeDbl((CIntSafe(txtMWT.Text) + _
                                                                    CIntSafe(txtWeightTolerance.Text)) / dblElementMass) + 1))
                        Else
                            intMultiplier = AbsMax(CIntSafe(frmFinderOptions.txtChargeMin), CIntSafe(frmFinderOptions.txtChargeMax))
                            txtMax(intIndex).Text = Trim(Str(CIntSafeDbl((CDblSafe(txtMWT.Text) + CIntSafeDbl(txtWeightTolerance.Text)) / dblElementMass) * intMultiplier + 1))
                        End If
                    ElseIf frmFinderOptions.cboSearchType.ListIndex = 0 Then
                        txtMax(intIndex).Text = Trim(Str(CIntSafeDbl((CDblSafe(txtPercentMaxWeight.Text) + _
                                         CIntSafeDbl(txtWeightTolerance.Text)) / dblElementMass) + 1))
                    Else
                        ' leave txtMax(intIndex).text unchanged
                    End If
                End If
            Next intIndex
        End If
    End If

End Sub

Private Sub UpdateStatus(intNumPotentialElementCount As Integer, intNumPointers As Integer, Optional blnInitialize As Boolean = False, Optional intMultipleSearchMax As Integer = 0)
    Static lngRecursiveCount As Long, strSpinStatus As String
    Static lngMaxRecursiveCount As Long
    
    If blnInitialize = True Then
        lngRecursiveCount = 0
        lngRecursiveFunctionCallCount = 0
        ' Calculate lngMaxRecursiveCount based on a combination function
        lngMaxRecursiveCount = Combinatorial(intNumPointers + intNumPotentialElementCount, intNumPotentialElementCount - 1) - Combinatorial(intNumPotentialElementCount + intNumPointers - 2, intNumPointers - 1)
        If intMultipleSearchMax > 0 Then
            ' Correct lngMaxRecursiveCount for searching for m/z values
            lngMaxRecursiveCount = lngMaxRecursiveCount * intMultipleSearchMax
        End If
        If intNumPotentialElementCount = 1 Then lngMaxRecursiveCount = 1
    Else
        lngRecursiveCount = lngRecursiveCount + 1
    End If
  
    If lngRecursiveCount <= lngMaxRecursiveCount Then
        lblPercentCompleted = LookupLanguageCaption(10120, "Searching") & " " & _
                              Format$(CIntSafeDbl(lngRecursiveCount / lngMaxRecursiveCount * 100), "##0") & _
                              LookupLanguageCaption(10110, "% Completed")
    Else
        Select Case Right(lblPercentCompleted.Caption, 2)
        Case "--": strSpinStatus = "\ "
        Case "\ ": strSpinStatus = "| "
        Case "|": strSpinStatus = "/ "
        Case "/ ": strSpinStatus = "--"
        Case Else: strSpinStatus = "--"
        End Select
        lblPercentCompleted = LookupLanguageCaption(10125, "Working") & " " & strSpinStatus
    End If
        
End Sub

Private Sub chkElements_Click(Index As Integer)
    Dim boolSetValue As Boolean, intIndex As Integer
    
    If gKeyPressAbortFormulaFinder = 1 Then
        gKeyPressAbortFormulaFinder = 3
    End If
    
    boolSetValue = cChkBox(chkElements(Index))

    ' If doing bounded search, show/hide min/max boxes
    If frmFinderOptions.cboSearchType.ListIndex = 1 Then
        ' Bounded search, so show/hide bounds boxes
        txtMin(Index).Visible = boolSetValue
        txtMax(Index).Visible = boolSetValue
        If boolSetValue = True Then
            lblMin.Visible = True
            lblMax.Visible = True
        End If
    End If
        
    ' No matter what, show/hide percent boxes if matching percents
    If optType(1).value = True Then
        txtPercent(Index).Visible = boolSetValue
        If boolSetValue = True Then
            lblPercent.Visible = True
        End If
    End If
    
    ' No matter what, for custom elements, show/hide weight box
    If Index >= 4 Then
        ' Custom elements boxes
        ' Display Atomic Wt. if box is checked
        txtWeight(Index).Visible = boolSetValue
        If boolSetValue = True Then lblWeight.Visible = True
    End If
    
    ' If hydrogen is present, then show verify hydrogens box
    If Index = 1 Then frmFinderOptions.chkVerifyHydrogens.Enabled = boolSetValue
            
    ' See if Min, Max, and Percent labels need to be removed
    If boolSetValue = False Then
        For intIndex = 0 To 9
            If cChkBox(chkElements(intIndex)) Then Exit For
        Next intIndex
        If intIndex = 10 Then
            lblMin.Visible = False
            lblMax.Visible = False
            lblPercent.Visible = False
            lblWeight.Visible = False
            cmdCalculate.Enabled = False
        Else
            ' See if Custom Weight label needs to be removed
            For intIndex = 4 To 9
                If cChkBox(chkElements(intIndex)) Then Exit For
            Next intIndex
            If intIndex = 10 Then lblWeight.Visible = False
        End If
    Else
        cmdCalculate.Enabled = True
    End If
    
    ' Determine position of lblWeight
    Call PositionLblWeight

    UpdateMinMax Index
    
End Sub

Private Sub chkPpmMode_Click()
    Dim temp As Double, ppm As Double
    
    If chkPpmMode.value = 0 Then
        ' now turning false, convert weight tolerance to regular
        ppm = CDblSafe(txtWeightTolerance.Text)
        If CDblSafe(txtMWT.Text) = 0 Then
            temp = 0.5
        Else
            temp = Abs(CDblSafe(txtMWT.Text) * ppm / 1000000#)
        End If
        txtWeightTolerance.Text = CStr(temp)
    Else
        ' now turning true, convert weight tolerance to ppm
        temp = CDblSafe(txtWeightTolerance.Text)
        ' Original formula was ppm = Abs((1 - (Val(txtMWT.Text) + temp) / Val(txtMWT.Text)) * 1000000#)
        ' Simplified using Derive to get:
        If Val(txtMWT.Text) = 0 Then
            ppm = 100
        Else
            ppm = Abs(1000000# * temp / CDblSafe(txtMWT.Text))
        End If
        txtWeightTolerance.Text = CStr(ppm)
    End If
    
End Sub

Private Sub cmdAbort_Click()
    gKeyPressAbortFormulaFinder = 3
End Sub

Private Sub cmdCalculate_Click()
    FormulaFinderCalculate
End Sub

Private Sub cmdCalculate_KeyDown(KeyCode As Integer, Shift As Integer)
    If gKeyPressAbortFormulaFinder = 1 Then
        If KeyCode >= 32 Then
            gKeyPressAbortFormulaFinder = 2
            KeyCode = 0
            Shift = 0
        End If
    End If

End Sub

Private Sub cmdCopy_Click()
    Dim X%, copytext$
    
    If lstResults.ListCount = 0 Then
        MsgBox LookupLanguageCaption(10550, "The results box is empty."), vbOKOnly, _
               LookupLanguageCaption(10555, "Nothing to Copy")
    Else
        Clipboard.Clear
        copytext$ = ""
        For X = 0 To lstResults.ListCount - 1
            ' MW not found, copy line to clipboard without any tabs
            copytext$ = copytext$ & vbCrLf & lstResults.List(X)
        Next X
        Clipboard.SetText copytext$, vbCFText
    End If
    
End Sub

Private Sub cmdCopyAsRTF_Click()
    If lstResults.ListCount = 0 Then
        MsgBox LookupLanguageCaption(10550, "The results box is empty."), vbOKOnly, _
               LookupLanguageCaption(10555, "Nothing to Copy")
    Else
        Clipboard.Clear
        
        Clipboard.SetText RemoveRTFHeightAdjust(rtfResults.TextRTF), vbCFRTF
    End If
    
End Sub

Private Sub cmdDisplayIsotopicAbundance_Click()
    DisplayIsoAbundanceForCurrent
End Sub

Private Sub cmdFinderOptions_Click()
   
    frmFinderOptions.Show
    
End Sub

Private Sub cmdOK_Click()
    HideFormShowMain Me
End Sub

Private Sub cmdPrint_Click()
    Dim X%, eResponse As VbMsgBoxResult

    If gKeyPressAbortFormulaFinder = 1 Then
        gKeyPressAbortFormulaFinder = 3
        Exit Sub
    End If

    If lstResults.ListCount = 0 Then
        MsgBox LookupLanguageCaption(10550, "The results box is empty."), vbOKOnly, _
               LookupLanguageCaption(10560, "Nothing to Print")
    Else
        ' Print the results
        eResponse = YesNoBox(LookupLanguageCaption(10565, "Are you sure you want to print the current result(s)?"), LookupLanguageCaption(10570, "Printing"))
        If eResponse = vbYes Then
            For X = 0 To lstResults.ListCount - 1
                Printer.Print lstResults.List(X)
            Next X
            Printer.EndDoc
        End If
    End If

End Sub

Private Sub Form_Activate()
    
    lstResults.FontName = objMwtWin.RtfFontName
    lstResults.FontSize = objMwtWin.RtfFontSize
    gKeyPressAbortFormulaFinder = 0
    
    PossiblyHideMainWindow

    ' Display correct weight mode phrase
    Select Case objMwtWin.GetElementMode
    Case 3: lblWeightMode.Caption = LookupLanguageCaption(3840, "(using integer isotopic weights)")
    Case 2: lblWeightMode.Caption = LookupLanguageCaption(3830, "(using isotopic elemental weights)")
    Case Else: lblWeightMode.Caption = LookupLanguageCaption(3820, "(using average atomic weights)")
    End Select

    LoadDynamicTextCaptions
End Sub

Private Sub Form_Click()
    If gKeyPressAbortFormulaFinder = 1 Then
        gKeyPressAbortFormulaFinder = 3
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If gKeyPressAbortFormulaFinder = 1 Then
        If KeyCode >= 32 Then
            gKeyPressAbortFormulaFinder = 2
            KeyCode = 0
            Shift = 0
        End If
    Else
        If KeyCode = vbKeyEscape Then cmdOK_Click
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 15 Then
        ' User pressed CTRL-O (finder options)
        cmdFinderOptions_Click
    End If

End Sub

Private Sub Form_Load()
    PositionFormControls
    
    ' Put window in center of screen
    SizeAndCenterWindow Me, cWindowTopCenter, 9100, 6300
        
    mnuRightClick.Visible = False
    
    ResizeForm True
    
    optType_Click 0
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If gKeyPressAbortFormulaFinder = 1 Then
        gKeyPressAbortFormulaFinder = 3
    End If
    
    QueryUnloadFormHandler Me, Cancel, UnloadMode
End Sub

Private Sub Form_Resize()
    ResizeForm
End Sub

Private Sub lblElement_Click()
    If gKeyPressAbortFormulaFinder = 1 Then
        gKeyPressAbortFormulaFinder = 3
    End If
End Sub

Private Sub lstResults_DblClick()
    ZoomLine lstResults.List(lstResults.ListIndex)
End Sub

Private Sub lstResults_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn
        lstResults_DblClick
    Case Else
        KeyAscii = 0
    End Select

End Sub

Private Sub mnuRightClickCopy_Click()
    CopyRoutinePrivate
End Sub

Private Sub mnuRightClickCut_Click()
    ' Can't cut from rtfResults, but will copy to the clipboard anyway
    CopyRoutinePrivate
End Sub

Private Sub mnuRightClickDelete_Click()
    ' Not used in frmFinder
End Sub

Private Sub mnuRightClickPaste_Click()
    ' Not used in frmFinder
End Sub

Private Sub mnuRightClickSelectAll_Click()
    rtfResults.SelStart = 0
    rtfResults.SelLength = Len(rtfResults.Text)
End Sub

Private Sub mnuRightClickUndo_Click()
    ' Not used in frmFinder
End Sub

Private Sub optType_Click(Index As Integer)
    SetWeightMatchingMode Index
End Sub

Private Sub rtfResults_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
    Case vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight
        ' Arrow keys are OK
    Case vbKeyPageUp, vbKeyPageDown, vbKeyHome, vbKeyEnd
        ' ok
    Case 16, 17, 18, vbKeyEscape, vbKeyF1
        ' Shift, Ctrl, Alt, Escape, & F1 keys
    Case Else
        If Shift And vbCtrlMask Then        ' And them in case alt or shift was also accidentally pressed
            If KeyCode = vbKeyC Then
                ' Allow Ctrl+C (copy), but not cut or paste
            ElseIf KeyCode = vbKeyA Then
                ' Allow Ctrl+A (Select All)
            ElseIf KeyCode = vbKeyD Then
                DisplayIsoAbundanceForCurrent True
                KeyCode = 0
                Shift = 0
            Else
                KeyCode = 0
                Shift = 0
            End If
        Else
            If Shift = vbAltMask Then
                ' Pressed Alt + some key; let it pass
                If KeyCode = vbKeyD Or KeyCode = vbKeyD + 32 Then
                    DisplayIsoAbundanceForCurrent True
                    KeyCode = 0
                    Shift = 0
                End If
            Else
                KeyCode = 0
                Shift = 0
            End If
        End If
    End Select
    
End Sub

Private Sub rtfResults_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        ShowPopupRightClickMenu Me, rtfResults, False, False
    End If
End Sub

Private Sub txtHits_Click()
    If gKeyPressAbortFormulaFinder = 1 Then
        gKeyPressAbortFormulaFinder = 3
    End If

End Sub

Private Sub txtHits_GotFocus()
    HighlightOnFocus txtHits

End Sub

Private Sub txtHits_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtHits, KeyAscii, True

End Sub

Private Sub txtMax_Click(Index As Integer)
    If gKeyPressAbortFormulaFinder = 1 Then
        gKeyPressAbortFormulaFinder = 3
    End If

End Sub

Private Sub txtMax_GotFocus(Index As Integer)
    HighlightOnFocus txtMax(Index)

End Sub

Private Sub txtMax_KeyPress(Index As Integer, KeyAscii As Integer)
    TextBoxKeyPressHandler txtMax(Index), KeyAscii, True
    
End Sub

Private Sub txtMin_Click(Index As Integer)
    If gKeyPressAbortFormulaFinder = 1 Then
        gKeyPressAbortFormulaFinder = 3
    End If

End Sub

Private Sub txtMin_GotFocus(Index As Integer)
    HighlightOnFocus txtMin(Index)

End Sub

Private Sub txtMin_KeyPress(Index As Integer, KeyAscii As Integer)
    TextBoxKeyPressHandler txtMin(Index), KeyAscii, True

End Sub

Private Sub txtMWT_Change()
    UpdateMinMax
End Sub

Private Sub txtMWT_Click()
    If gKeyPressAbortFormulaFinder = 1 Then
        gKeyPressAbortFormulaFinder = 3
    End If
End Sub

Private Sub txtMWT_GotFocus()
    HighlightOnFocus txtMWT
End Sub

Private Sub txtMWT_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtMWT, KeyAscii, True, True

End Sub

Private Sub txtPercent_Click(Index As Integer)
    If gKeyPressAbortFormulaFinder = 1 Then
        gKeyPressAbortFormulaFinder = 3
    End If

End Sub

Private Sub txtPercent_GotFocus(Index As Integer)
    HighlightOnFocus txtPercent(Index)
    
End Sub

Private Sub txtPercent_KeyPress(Index As Integer, KeyAscii As Integer)
    TextBoxKeyPressHandler txtPercent(Index), KeyAscii, True, True

End Sub

Private Sub txtPercentMaxWeight_GotFocus()
    HighlightOnFocus txtPercentMaxWeight

End Sub

Private Sub txtPercentMaxWeight_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtPercentMaxWeight, KeyAscii, True, True

End Sub

Private Sub txtPercentTolerance_Click()
    If gKeyPressAbortFormulaFinder = 1 Then
        gKeyPressAbortFormulaFinder = 3
    End If

End Sub

Private Sub txtPercentTolerance_GotFocus()
    HighlightOnFocus txtPercentTolerance

End Sub

Private Sub txtPercentTolerance_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtPercentTolerance, KeyAscii, True, True

End Sub

Private Sub txtWeight_Click(Index As Integer)
    If gKeyPressAbortFormulaFinder = 1 Then
        gKeyPressAbortFormulaFinder = 3
    End If

End Sub
Private Sub txtWeight_GotFocus(Index As Integer)
    HighlightOnFocus txtWeight(Index)

End Sub

Private Sub txtWeight_KeyPress(Index As Integer, KeyAscii As Integer)
    TextBoxKeyPressHandler txtWeight(Index), KeyAscii, True, True, False, True, True, True

End Sub

Private Sub txtWeightTolerance_Click()
    If gKeyPressAbortFormulaFinder = 1 Then
        gKeyPressAbortFormulaFinder = 3
    End If

End Sub

Private Sub txtWeightTolerance_GotFocus()
    HighlightOnFocus txtWeightTolerance

End Sub

Private Sub txtWeightTolerance_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtWeightTolerance, KeyAscii, True, True

End Sub

