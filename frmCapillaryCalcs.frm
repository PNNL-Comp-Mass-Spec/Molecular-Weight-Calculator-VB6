VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmCapillaryCalcs 
   Caption         =   "Capillary Flow and Mass Rate Calculations"
   ClientHeight    =   6600
   ClientLeft      =   540
   ClientTop       =   1680
   ClientWidth     =   10770
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
   HelpContextID   =   3070
   Icon            =   "frmCapillaryCalcs.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6600
   ScaleWidth      =   10770
   Tag             =   "7000"
   Begin VB.CommandButton cmdComputeViscosity 
      Caption         =   "Compute Water/MeCN Viscosity"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   240
      TabIndex        =   30
      Tag             =   "7740"
      Top             =   2880
      Width           =   2715
   End
   Begin VB.Frame fraBroadening 
      Caption         =   "Extra Column Broadening Calculations"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2880
      Left            =   120
      TabIndex        =   48
      Tag             =   "7350"
      Top             =   3240
      Visible         =   0   'False
      Width           =   10455
      Begin VB.ComboBox cboCapValue 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   18
         Left            =   9240
         Style           =   2  'Dropdown List
         TabIndex        =   94
         Tag             =   "7070"
         Top             =   2040
         Width           =   1095
      End
      Begin VB.ComboBox cboCapValue 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   17
         Left            =   3720
         Style           =   2  'Dropdown List
         TabIndex        =   93
         Tag             =   "7070"
         Top             =   2040
         Width           =   1095
      End
      Begin VB.ComboBox cboCapValue 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   15
         Left            =   3720
         Style           =   2  'Dropdown List
         TabIndex        =   57
         Tag             =   "7035"
         Top             =   1320
         Width           =   1815
      End
      Begin VB.ComboBox cboCapValue 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   16
         Left            =   3720
         Style           =   2  'Dropdown List
         TabIndex        =   60
         Tag             =   "7035"
         Top             =   1680
         Width           =   1815
      End
      Begin VB.TextBox txtCapValue 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   12
         Left            =   2520
         TabIndex        =   82
         Text            =   "0.000001"
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton cmdViewBroadeningEquations 
         Caption         =   "View Equations"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   81
         Tag             =   "7730"
         Top             =   2400
         Width           =   1815
      End
      Begin VB.TextBox txtCapValue 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   16
         Left            =   8160
         TabIndex        =   71
         Text            =   "1"
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox txtCapValue 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   15
         Left            =   2520
         TabIndex        =   62
         Text            =   "1"
         Top             =   2040
         Width           =   1095
      End
      Begin RichTextLib.RichTextBox rtfBdTemporalVarianceUnit 
         Height          =   255
         Left            =   9360
         TabIndex        =   69
         TabStop         =   0   'False
         Tag             =   "7550"
         Top             =   960
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393217
         BackColor       =   -2147483633
         BorderStyle     =   0
         Appearance      =   0
         TextRTF         =   $"frmCapillaryCalcs.frx":08CA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox txtCapValue 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   14
         Left            =   2520
         TabIndex        =   59
         Text            =   "1"
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox txtCapValue 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   13
         Left            =   2520
         TabIndex        =   56
         Text            =   "1"
         Top             =   1320
         Width           =   1095
      End
      Begin VB.ComboBox cboCapValue 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   14
         Left            =   3720
         Style           =   2  'Dropdown List
         TabIndex        =   52
         Tag             =   "7060"
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox txtCapValue 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   11
         Left            =   2520
         TabIndex        =   51
         Text            =   "1"
         Top             =   360
         Width           =   1095
      End
      Begin VB.CheckBox chkBdLinkLinearVelocity 
         Caption         =   "&Link to Above"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   50
         Tag             =   "7800"
         Top             =   600
         Width           =   2055
      End
      Begin RichTextLib.RichTextBox rtfBdAdditionalVarianceUnit 
         Height          =   255
         Left            =   9360
         TabIndex        =   72
         TabStop         =   0   'False
         Tag             =   "7550"
         Top             =   1320
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393217
         BackColor       =   -2147483633
         BorderStyle     =   0
         Appearance      =   0
         TextRTF         =   $"frmCapillaryCalcs.frx":093D
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox rtfDiffusionCoefficient 
         Height          =   255
         Left            =   3720
         TabIndex        =   54
         TabStop         =   0   'False
         Tag             =   "7520"
         Top             =   960
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         _Version        =   393217
         BackColor       =   -2147483633
         BorderStyle     =   0
         Enabled         =   0   'False
         Appearance      =   0
         TextRTF         =   $"frmCapillaryCalcs.frx":09B0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblDiffusionCoefficient 
         Caption         =   "Diffusion Coefficient"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   53
         Tag             =   "7360"
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label lblOptimumLinearVelocityBasis 
         Caption         =   "(for 5 um particles)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5880
         TabIndex        =   66
         Tag             =   "7410"
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label lblOptimumLinearVelocity 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8160
         TabIndex        =   64
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblOptimumLinearVelocityUnit 
         Caption         =   "cm/sec"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9360
         TabIndex        =   65
         Tag             =   "7540"
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblOptimumLinearVelocityLabel 
         Caption         =   "Optimum Linear Velocity"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5760
         TabIndex        =   63
         Tag             =   "7400"
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label lblBdPercentVarianceIncrease 
         Caption         =   "0%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8160
         TabIndex        =   76
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label lblBdPercentVarianceIncreaseLabel 
         Caption         =   "Percent Increase"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5760
         TabIndex        =   75
         Tag             =   "7450"
         Top             =   2400
         Width           =   2295
      End
      Begin VB.Label lblBdResultantPeakWidthLabel 
         Caption         =   "Resulting Peak Width"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5760
         TabIndex        =   73
         Tag             =   "7440"
         Top             =   2040
         Width           =   2295
      End
      Begin VB.Label lblBdResultantPeakWidth 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8160
         TabIndex        =   74
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label lblBdAdditionalVarianceLabel 
         Caption         =   "Additional Variance"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5760
         TabIndex        =   70
         Tag             =   "7430"
         Top             =   1320
         Width           =   2295
      End
      Begin VB.Label lblBdInitialPeakWidth 
         Caption         =   "Initial Peak Width (at base)"
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
         Left            =   240
         TabIndex        =   61
         Tag             =   "7390"
         Top             =   2040
         Width           =   2175
      End
      Begin VB.Label lblBdTemporalVarianceLabel 
         Caption         =   "Temporal Variance"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5760
         TabIndex        =   67
         Tag             =   "7420"
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label lblBdTemporalVariance 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8160
         TabIndex        =   68
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lblBdOpenTubeLength 
         Caption         =   "Open Tube Length"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   55
         Tag             =   "7370"
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label lblBdOpenTubeID 
         Caption         =   "Open Tube I.D."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   58
         Tag             =   "7380"
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label lblBdLinearVelocity 
         Caption         =   "Linear Velocity"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   49
         Tag             =   "7260"
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.ComboBox cboCapValue 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   4
      Left            =   3840
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Tag             =   "7035"
      Top             =   2520
      Width           =   1815
   End
   Begin VB.ComboBox cboCapValue 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   3840
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Tag             =   "7035"
      Top             =   1440
      Width           =   1815
   End
   Begin VB.ComboBox cboCapValue 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   3840
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Tag             =   "7035"
      Top             =   1800
      Width           =   1815
   End
   Begin VB.TextBox txtCapValue 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   6
      Left            =   8280
      TabIndex        =   80
      Text            =   "1"
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton cmdViewEquations 
      Caption         =   "&View Explanatory Equations"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   7440
      TabIndex        =   79
      Tag             =   "7710"
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton cmdShowPeakBroadening 
      Caption         =   "Show/Hide &Peak Broadening Calculations"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   5160
      TabIndex        =   78
      Tag             =   "7700"
      Top             =   120
      Width           =   2115
   End
   Begin VB.ComboBox cboCapValue 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   8
      Left            =   9480
      Style           =   2  'Dropdown List
      TabIndex        =   27
      Tag             =   "7080"
      Top             =   2160
      Width           =   1095
   End
   Begin VB.ComboBox cboCapValue 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   7
      Left            =   9480
      Style           =   2  'Dropdown List
      TabIndex        =   24
      Tag             =   "7070"
      Top             =   1800
      Width           =   1095
   End
   Begin VB.ComboBox cboCapValue 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   6
      Left            =   9480
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Tag             =   "7060"
      Top             =   1440
      Width           =   1095
   End
   Begin VB.ComboBox cboCapValue 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   5
      Left            =   9480
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Tag             =   "7050"
      Top             =   1080
      Width           =   1095
   End
   Begin VB.ComboBox cboCapValue 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   3840
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Tag             =   "7040"
      Top             =   2160
      Width           =   1815
   End
   Begin VB.ComboBox cboCapValue 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   3840
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Tag             =   "7030"
      Top             =   1080
      Width           =   1815
   End
   Begin VB.TextBox txtCapValue 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   7
      Left            =   9480
      TabIndex        =   29
      Text            =   ".4"
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox txtCapValue 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   8280
      TabIndex        =   18
      Text            =   "1"
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox txtCapValue 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   2640
      TabIndex        =   15
      Text            =   "1"
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox txtCapValue 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   2640
      TabIndex        =   12
      Text            =   "1"
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox txtCapValue 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   2640
      TabIndex        =   9
      Text            =   "1"
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox txtCapValue 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   2640
      TabIndex        =   6
      Text            =   "1"
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox txtCapValue 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   2640
      TabIndex        =   3
      Text            =   "1"
      Top             =   1080
      Width           =   1095
   End
   Begin VB.ComboBox cboComputationType 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Tag             =   "7020"
      Top             =   520
      Width           =   4575
   End
   Begin VB.Frame fraMassRate 
      Caption         =   "Mass Rate Calculations"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   31
      Tag             =   "7300"
      Top             =   3240
      Width           =   10455
      Begin VB.ComboBox cboCapValue 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   13
         Left            =   9120
         Style           =   2  'Dropdown List
         TabIndex        =   47
         Tag             =   "7110"
         Top             =   1320
         Width           =   1215
      End
      Begin VB.ComboBox cboCapValue 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   12
         Left            =   9120
         Style           =   2  'Dropdown List
         TabIndex        =   44
         Tag             =   "7100"
         Top             =   360
         Width           =   1215
      End
      Begin VB.ComboBox cboCapValue 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   11
         Left            =   3720
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Tag             =   "7070"
         Top             =   1320
         Width           =   1695
      End
      Begin VB.ComboBox cboCapValue 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   10
         Left            =   3720
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Tag             =   "7050"
         Top             =   720
         Width           =   1695
      End
      Begin VB.ComboBox cboCapValue 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   9
         Left            =   3720
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Tag             =   "7090"
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox txtCapValue 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   10
         Left            =   2520
         TabIndex        =   40
         Text            =   "1"
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox txtCapValue 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   9
         Left            =   2520
         TabIndex        =   37
         Text            =   "1"
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtCapValue 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   8
         Left            =   2520
         TabIndex        =   33
         Text            =   "1"
         Top             =   360
         Width           =   1095
      End
      Begin VB.CheckBox chkMassRateLinkFlowRate 
         Caption         =   "&Link to Above"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   36
         Tag             =   "7800"
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label lblMolesInjectedLabel 
         Caption         =   "Moles Injected"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5760
         TabIndex        =   45
         Tag             =   "7340"
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label lblMolesInjected 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8040
         TabIndex        =   46
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label lblMassFlowRateLabel 
         Caption         =   "Mass Flow Rate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5760
         TabIndex        =   42
         Tag             =   "7330"
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label lblMassFlowRate 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8040
         TabIndex        =   43
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblMassRateInjectionTime 
         Caption         =   "Injection Time"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   39
         Tag             =   "7320"
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label lblMassRateVolFlowRate 
         Caption         =   "Volumetric Flow Rate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Tag             =   "7250"
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label lblMassRateConcentration 
         Caption         =   "Sample Concentration"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Tag             =   "7310"
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.ComboBox cboCapillaryType 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Tag             =   "7010"
      ToolTipText     =   "Toggle between open and packed capillaries"
      Top             =   120
      Width           =   2655
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "Cl&ose"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9360
      TabIndex        =   77
      Tag             =   "4000"
      Top             =   240
      Width           =   1155
   End
   Begin VB.Frame fraWeightSource 
      BorderStyle     =   0  'None
      Caption         =   "Weight Source"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   240
      TabIndex        =   83
      Top             =   5040
      Width           =   5655
      Begin VB.OptionButton optWeightSource 
         Caption         =   "&Use mass of compound in current formula"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         HelpContextID   =   4010
         Index           =   0
         Left            =   0
         TabIndex        =   86
         Tag             =   "7750"
         Top             =   0
         Value           =   -1  'True
         Width           =   5475
      End
      Begin VB.OptionButton optWeightSource 
         Caption         =   "&Enter custom numerical mass"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         HelpContextID   =   4010
         Index           =   1
         Left            =   0
         TabIndex        =   85
         Tag             =   "7760"
         Top             =   240
         Width           =   5475
      End
      Begin VB.TextBox txtCustomMass 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3360
         TabIndex        =   84
         Tag             =   "7600"
         Text            =   "100"
         ToolTipText     =   "Enter custom numerical mass for use in computations"
         Top             =   1200
         Width           =   1575
      End
      Begin RichTextLib.RichTextBox rtfCurrentFormula 
         Height          =   495
         Left            =   2040
         TabIndex        =   87
         TabStop         =   0   'False
         Top             =   600
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   873
         _Version        =   393217
         Enabled         =   0   'False
         MultiLine       =   0   'False
         Appearance      =   0
         TextRTF         =   $"frmCapillaryCalcs.frx":0A26
      End
      Begin VB.Label lblMWTValue 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   92
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label lblCustomMassUnits 
         Caption         =   "g/mole"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5040
         TabIndex        =   91
         Tag             =   "7570"
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lblMWT 
         Caption         =   "MW ="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   90
         Tag             =   "7470"
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label lblFormula 
         Caption         =   "Current Formula is"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   89
         Tag             =   "7460"
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label lblCustomMass 
         Caption         =   "Custom Mass:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   88
         Tag             =   "7480"
         Top             =   1200
         Width           =   2175
      End
   End
   Begin VB.Label lblVolumeLabel 
      Caption         =   "Column Volume"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5880
      TabIndex        =   25
      Tag             =   "7280"
      Top             =   2160
      Width           =   2295
   End
   Begin VB.Label lblVolume 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8280
      TabIndex        =   26
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label lblDeadTimeLabel 
      Caption         =   "Column Dead Time"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5880
      TabIndex        =   23
      Tag             =   "7270"
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Label lblLinearVelocity 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8280
      TabIndex        =   21
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label lblLinearVelocityLabel 
      Caption         =   "Linear Velocity"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5880
      TabIndex        =   20
      Tag             =   "7260"
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label lblFlowRateLabel 
      Caption         =   "Volumetric Flow Rate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5880
      TabIndex        =   17
      Tag             =   "7250"
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label lblPorosity 
      Caption         =   "Interparticle Porosity (epsilon)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5880
      TabIndex        =   28
      Tag             =   "7290"
      Top             =   2520
      Width           =   3495
   End
   Begin VB.Label lblParticleDiameter 
      Caption         =   "Particle Diameter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Tag             =   "7240"
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Label lblViscosity 
      Caption         =   "Solvent Viscosity"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Tag             =   "7230"
      Top             =   2160
      Width           =   2295
   End
   Begin VB.Label lblColumnID 
      Caption         =   "Column Inner Diameter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Tag             =   "7220"
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label lblLength 
      Caption         =   "Column Length"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Tag             =   "7210"
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label lblPressure 
      Caption         =   "Back Pressure"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Tag             =   "7200"
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuLoadCapValues 
         Caption         =   "&Load Values"
      End
      Begin VB.Menu mnuSaveCapValues 
         Caption         =   "&Save Values"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuCapillaryFlowHelp 
         Caption         =   "&Capillary Flow Calculations"
      End
   End
End
Attribute VB_Name = "frmCapillaryCalcs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Subs ResizeCapillaryCalcsForm and ShowHideWeightSource in frmCapillaryCalcs
'   use cConcentrationUnitsFirstWeightIndex to determine whether
'   weight-based (mg/mL) or mole-based (Molar) mass units are being used
Private Const cConcentrationUnitsFirstWeightIndex = 7

Private Enum cptComputationTypeConstants
    cptBackPressure = 0
    cptColumnLength = 1
    cptColumnID = 2
    cptVolFlowRate = 3
    cptVolFlowRateUsingDeadTime = 4
End Enum

' Constants for Combo Boxes on Capillay Calcs form
Private Enum caCapillaryActionConstants
     caFindBackPressure = 0
     caFindColumnLength = 1
     caFindInnerDiameter = 2
     caFindVolFlowRateUsingPressure = 3
     caFindVolFlowUsingDeadTime = 4
End Enum

Private Const CAPILLARY_CALCS_FORM_INITIAL_WIDTH = 10800
Private Const CAPILLARY_CALCS_FORM_INITIAL_HEIGHT = 6000

' Note: Other forms access this variable
Public eMassMode As mmcMassModeConstants

' Form-wide variables
Private mDefaultCapValuesLoaded As Boolean
Private mUpdatingCapValues As Boolean
Private eColumnIDUnitsIndexSaved As ulnUnitsLengthConstants

Private objCapillaryFlow As New MWCapillaryFlowClass

' Purpose: Examines current computation type and checks to see if user (or code)
'          is attempting to modify a text box in which a result will be placed
'          If they (or the code) is, then do not allow any of the auto-update cascade events to fire
Private Function CheckDoNotModify(eTextBoxID As cctCapCalcTextBoxIDConstants) As Boolean
    
    Dim blnDoNotModify As Boolean
    
    Select Case cboComputationType.ListIndex
    Case caFindBackPressure
        If eTextBoxID = cctPressure Then blnDoNotModify = True
    Case caFindColumnLength
        If eTextBoxID = cctColumnLength Then blnDoNotModify = True
    Case caFindInnerDiameter
        If eTextBoxID = cctColumnID Then blnDoNotModify = True
    Case caFindVolFlowRateUsingPressure
        If eTextBoxID = cctFlowRate Then blnDoNotModify = True
    Case caFindVolFlowUsingDeadTime
        ' Finding Vol. Flow rate and pressure using dead time; do not auto-compute on change
        If eTextBoxID = cctPressure Or eTextBoxID = cctFlowRate Then blnDoNotModify = True
    End Select

    CheckDoNotModify = blnDoNotModify
End Function

' Purpose: Copy flow info from txtFlowRate to txtMassRateVolFlowRate
Private Sub CopyFlowRate()
    ' Use a static variable to prevent infinite loops from occurring
    Static blnCopyingFlowRate As Boolean
    
    If Not blnCopyingFlowRate Then
        If cChkBox(chkMassRateLinkFlowRate) Then
            blnCopyingFlowRate = True
            
            txtCapValue(cctMassRateVolFlowRate).Text = txtCapValue(cctFlowRate).Text
            cboCapValue(cccMassRateVolFlowRateUnits).ListIndex = cboCapValue(cccFlowRateUnits).ListIndex
            
            blnCopyingFlowRate = False
        End If
    End If
End Sub

' Purpose: Copy velocity info from lblLinearVelocity to txtBdLinearVelocity
Private Sub CopyLinearVelocity()
    txtCapValue(cctBdLinearVelocity).Text = lblLinearVelocity.Caption
    cboCapValue(cccBdLinearVelocityUnits).ListIndex = cboCapValue(cccLinearVelocityUnits).ListIndex
End Sub

' Purpose: Put appropriate labels in RTF boxes (labels change based on language)
Private Sub FillRTFBoxes()
    Dim strRtfText As String, strRtfPreString As String, strRtfTwoSquared As String
    Dim strWorkingCaption As String, strNumerator As String, strDenominator As String
    Dim intCaretLoc As Integer
    
    ' Put text in RTF text boxes
    strRtfPreString = "{\rtf1\ansi\ansicpg1252\deff0\deftab720{\fonttbl{\f0\fswiss MS Sans Serif;}{\f1\froman\fcharset2 Symbol;}{\f2\froman Times New Roman;}{\f3\fswiss MS Sans Serif;}{\f4\froman Times New Roman;}}"
    strRtfPreString = strRtfPreString & "{\colortbl\red0\green0\blue0;\red255\green255\blue255;}"
    strRtfPreString = strRtfPreString & "\deflang1033\pard\plain\f0\fs16 "
    strRtfTwoSquared = "\plain\f0\fs16\up3 2\plain\f0\fs16 "
    
    strWorkingCaption = LookupLanguageCaption(7520, "cm^2/sec")
    intCaretLoc = InStr(strWorkingCaption, "^")
    If InStr(strWorkingCaption, "^") Then
        strNumerator = Left(strWorkingCaption, intCaretLoc - 1)
        strDenominator = Mid(strWorkingCaption, intCaretLoc + 2)
    Else
        strNumerator = "cm"
        strDenominator = "/sec"
    End If
    
    strRtfText = strRtfPreString & strNumerator & strRtfTwoSquared & strDenominator
    strRtfText = strRtfText & "\par }"
        rtfDiffusionCoefficient.TextRTF = strRtfText

    strWorkingCaption = LookupLanguageCaption(7550, "sec^2")
    intCaretLoc = InStr(strWorkingCaption, "^")
    If InStr(strWorkingCaption, "^") Then
        strNumerator = Left(strWorkingCaption, intCaretLoc - 1)
    Else
        strNumerator = "sec"
    End If
    
    strRtfText = strRtfPreString & strNumerator & strRtfTwoSquared & "\par }"
        rtfBdTemporalVarianceUnit.TextRTF = strRtfText
        rtfBdAdditionalVarianceUnit.TextRTF = strRtfText
    

End Sub

Private Sub FindDesiredValue()
    Dim dblValue As Double, dblNewBackPressure As Double
    
    Select Case cboComputationType.ListIndex
    Case cptBackPressure
         dblValue = objCapillaryFlow.ComputeBackPressure(cboCapValue(cccPressureUnits).ListIndex)
        FormatTextBox txtCapValue(cctPressure), dblValue
        
    Case cptColumnLength
        dblValue = objCapillaryFlow.ComputeColumnLength(cboCapValue(cccColumnLengthUnits).ListIndex)
        FormatTextBox txtCapValue(cctColumnLength), dblValue
        
    Case cptColumnID
        dblValue = objCapillaryFlow.ComputeColumnID(cboCapValue(cccColumnIDUnits).ListIndex)
        FormatTextBox txtCapValue(cctColumnID), dblValue
        
    Case cptVolFlowRateUsingDeadTime
        dblValue = objCapillaryFlow.ComputeVolFlowRateUsingDeadTime(cboCapValue(cccFlowRateUnits).ListIndex, dblNewBackPressure, cboCapValue(cccPressureUnits).ListIndex)
        FormatTextBox txtCapValue(cctFlowRate), dblValue
        FormatTextBox txtCapValue(cctPressure), dblNewBackPressure
        
    Case Else       ' Includes cptVolFlowRate
        dblValue = objCapillaryFlow.ComputeVolFlowRate(cboCapValue(cccFlowRateUnits).ListIndex)
        FormatTextBox txtCapValue(cctFlowRate), dblValue
    End Select
    
    ' Update Linear Velocity (it is auto-computed by any of the above calls to objCapillaryFlow)
    dblValue = objCapillaryFlow.GetLinearVelocity(cboCapValue(cccLinearVelocityUnits).ListIndex)
    FormatLabel lblLinearVelocity, dblValue
    
    ' Update the dead time, unless we computed the flow rate using the dead time
    If cboComputationType.ListIndex <> cptVolFlowRateUsingDeadTime Then
        dblValue = objCapillaryFlow.GetDeadTime(cboCapValue(cccDeadTimeUnits).ListIndex)
        FormatTextBox txtCapValue(cctDeadTime), dblValue
    End If
    
    ' Update the column volume
    dblValue = objCapillaryFlow.GetColumnVolume(cboCapValue(cccVolumeUnits).ListIndex)
    FormatLabel lblVolume, dblValue
    
End Sub

Private Sub FindMassRate()
    Dim dblValue As Double
    
    If Not mDefaultCapValuesLoaded Then Exit Sub

    ' Note: This automatically calls ComputeMassRateValues which computes
    '        the Mass Flow Rate and Moles Injected
    objCapillaryFlow.SetMassRateSampleMass GetWorkingMass()

    dblValue = objCapillaryFlow.GetMassFlowRate(cboCapValue(cccMassFlowRateUnits).ListIndex)
    FormatLabel lblMassFlowRate, dblValue
    
    dblValue = objCapillaryFlow.GetMassRateMolesInjected(cboCapValue(cccMassRateMolesInjectedUnits).ListIndex)
    FormatLabel lblMolesInjected, dblValue

End Sub

Private Sub FindBroadening()
    
    Dim dblInitialPeakWidthInSec As Double, dblResultantPeakWidthInSec As Double
    Dim dblPercentIncrease As Double
    Dim dblValue As Double
    
    ' Note: This will calculate the temporal variance, in addition to the resultant peak width
    dblValue = objCapillaryFlow.ComputeExtraColumnBroadeningResultantPeakWidth(cboCapValue(cccBdResultantPeakWidthUnits).ListIndex)
    FormatLabel lblBdResultantPeakWidth, dblValue
    
    dblValue = objCapillaryFlow.GetExtraColumnBroadeningTemporalVarianceInSquareSeconds()
    FormatLabel lblBdTemporalVariance, dblValue
    
    ' The peak widths were computed above
    ' Simply retrieve them, using seconds as the unit, and compute the percent increase in the peak width
    dblInitialPeakWidthInSec = objCapillaryFlow.GetExtraColumnBroadeningInitialPeakWidthAtBase(utmSeconds)
    dblResultantPeakWidthInSec = objCapillaryFlow.GetExtraColumnBroadeningResultantPeakWidth(utmSeconds)
    
    If dblInitialPeakWidthInSec <> 0 Then
        dblPercentIncrease = (dblResultantPeakWidthInSec - dblInitialPeakWidthInSec) / dblInitialPeakWidthInSec * 100
    Else
        dblPercentIncrease = 0
    End If
    
    lblBdPercentVarianceIncrease.Caption = Format(dblPercentIncrease / 100, "##0.0%")
    
End Sub

Private Sub FindOptimumLinearVelocity()
    ' ToDo: Make this update the linear velocity
    
    Dim dblOptimumLinearVelocity As Double, strCaptionWork As String
    Dim intNumStartLoc As Integer, intNumEndLoc As Integer, intIndex As Integer
    
    dblOptimumLinearVelocity = objCapillaryFlow.ComputeOptimumLinearVelocityUsingParticleDiamAndDiffusionCoeff(cboCapValue(cccLinearVelocityUnits).ListIndex)

    FormatLabel lblOptimumLinearVelocity, dblOptimumLinearVelocity
    lblOptimumLinearVelocityUnit = cboCapValue(cccLinearVelocityUnits).Text
    
    strCaptionWork = LookupLanguageCaption(7410, "(for 5 um particles)")
    intNumEndLoc = -1
    intNumStartLoc = -1
    For intIndex = 1 To Len(strCaptionWork)
        If intNumStartLoc <= 0 Then
            If IsNumeric(Mid(strCaptionWork, intIndex, 1)) Then
                intNumStartLoc = intIndex
            End If
        Else
            If Not IsNumeric(Mid(strCaptionWork, intIndex, 1)) Then
                intNumEndLoc = intIndex - 1
                Exit For
            End If
        End If
    Next intIndex
    
    With txtCapValue(cctParticleDiamter)
        If intNumEndLoc > 0 And intNumStartLoc > 0 Then
            strCaptionWork = Left(strCaptionWork, intNumStartLoc - 1) & Trim(.Text) & _
                              Mid(strCaptionWork, intNumEndLoc + 1)
            lblOptimumLinearVelocityBasis.Caption = strCaptionWork
        Else
            lblOptimumLinearVelocityBasis.Caption = "(for " & Trim(.Text) & " um particles)"
        End If
    End With
    
End Sub

' Purpose: Get working mass from lblMwtValue or txtCustomMass
Private Function GetWorkingMass() As Double
    If eMassMode = 0 Then
        GetWorkingMass = CDblSafe(lblMWTValue.Caption)
    Else
        GetWorkingMass = CDblSafe(txtCustomMass.Text)
    End If
End Function

' Purpose: Change background color of text boxes to indicate which is being found
Private Sub HighlightTargetBoxes()
    Dim ctlThisControl As Control, strControlType As String
    Dim lngHighlightedColor As Long
    
    For Each ctlThisControl In Me.Controls
        strControlType = TypeName(ctlThisControl)
        If strControlType = "TextBox" Then
            ctlThisControl.BackColor = QBColor(COLOR_WHITE)
        End If
    Next

    ' 7, 10, and 14 are good colors for QBColor()
    lngHighlightedColor = QBColor(COLOR_COMPUTEDQUANTITY)
    Select Case cboComputationType.ListIndex
    Case caFindBackPressure
        txtCapValue(cctPressure).BackColor = lngHighlightedColor
    Case caFindColumnLength
        txtCapValue(cctColumnLength).BackColor = lngHighlightedColor
    Case caFindInnerDiameter
        txtCapValue(cctColumnID).BackColor = lngHighlightedColor
    Case caFindVolFlowRateUsingPressure
        txtCapValue(cctFlowRate).BackColor = lngHighlightedColor
    Case caFindVolFlowUsingDeadTime
        ' Finding Vol. Flow rate and pressure using dead time; do not auto-compute on change
        txtCapValue(cctPressure).BackColor = lngHighlightedColor
        txtCapValue(cctFlowRate).BackColor = lngHighlightedColor
    End Select
    
    If cboComputationType.ListIndex <> caFindVolFlowUsingDeadTime Then
        txtCapValue(cctDeadTime).BackColor = lngHighlightedColor
    End If
    
End Sub

' Purpose: Load the default capillary form values
Private Sub LoadCapFormValues()
    Dim intIndex As Integer
    
On Error GoTo LoadCapFormValuesErrorHandler
    
    If Not mDefaultCapValuesLoaded Then Exit Sub

    mUpdatingCapValues = True
    
    cboComputationType.ListIndex = gCapFlowComputationTypeSave
    chkMassRateLinkFlowRate.value = gCapFlowLinkMassRateFlowRateSave
    chkBdLinkLinearVelocity.value = gCapFlowLinkBdLinearVelocitySave
    
    If gCapFlowShowPeakBroadeningSave = 1 And fraBroadening.Visible = False Then
        ShowHideBroadeningFrame
    End If
    
    If cboCapillaryType.ListIndex = ctOpenTubularCapillary Then
        ' Open Capillary
        For intIndex = 0 To CapTextBoxMaxIndex
            txtCapValue(intIndex).Text = Trim(CStr(OpenCapVals.TextValues(intIndex)))
        Next intIndex
        
        For intIndex = 0 To CapComboBoxMaxIndex
            If OpenCapVals.ComboValues(intIndex) < cboCapValue(intIndex).ListCount Then
                cboCapValue(intIndex).ListIndex = OpenCapVals.ComboValues(intIndex)
            End If
        Next intIndex
    Else
        ' Packed Capillary
        For intIndex = 0 To CapTextBoxMaxIndex
            txtCapValue(intIndex).Text = Trim(CStr(PackedCapVals.TextValues(intIndex)))
        Next intIndex
        
        For intIndex = 0 To CapComboBoxMaxIndex
            If PackedCapVals.ComboValues(intIndex) < cboCapValue(intIndex).ListCount Then
                cboCapValue(intIndex).ListIndex = PackedCapVals.ComboValues(intIndex)
            End If
        Next intIndex

    End If
    
    eColumnIDUnitsIndexSaved = cboCapValue(cccColumnIDUnits).ListIndex
    mUpdatingCapValues = False
    
    Exit Sub
    
LoadCapFormValuesErrorHandler:
    Debug.Print "Error in LoadCapFormValues"
    Debug.Assert False
    Resume Next
End Sub

' Purpose: Populate the combo boxes with the valid choices (language specific)
Private Sub PopulateComboBoxes()
    Dim intIndex As Integer
    
    ' Load Amount types in the Combo Boxes
    
    PopulateComboBox cboCapillaryType, True, "Open Tubular Capillary|Packed Capillary", 0   '7010
    PopulateComboBox cboComputationType, True, "Find Back Pressure|Find Column Length|Find Inner Diameter|Find Volumetric Flow rate|Find Flow Rate using Dead Time", 3   '7020
    
    PopulateComboBox cboCapValue(cccPressureUnits), True, "psi|Pascals|kiloPascals|Atmospheres|Bar|Torr (mm Hg)|dynes/cm^2", 0   '7030
    
    PopulateComboBox cboCapValue(cccColumnLengthUnits), True, "m|cm|mm|um|inches", 1    '7035
    PopulateComboBox cboCapValue(cccColumnIDUnits), True, "m|cm|mm|um|inches", 3        '7035
    
    PopulateComboBox cboCapValue(cccViscosityUnits), True, "Poise [g/(cm-sec)]|centiPoise", 0   '7040
    
    With cboCapValue(cccParticleDiameterUnits)    '7035 Dup
        For intIndex = 0 To cboCapValue(cccColumnIDUnits).ListCount - 1
            .AddItem cboCapValue(cccColumnIDUnits).List(intIndex)
        Next intIndex
        .ListIndex = cboCapValue(cccColumnIDUnits).ListIndex
    End With
    
    PopulateComboBox cboCapValue(cccFlowRateUnits), True, "mL/min|uL/min|nL/min", 2   '7050
    PopulateComboBox cboCapValue(cccLinearVelocityUnits), True, "cm/hr|mm/hr|cm/min|mm/min|cm/sec|mm/sec", 4   '7060
    PopulateComboBox cboCapValue(cccDeadTimeUnits), True, "hours|minutes|seconds", 2   '7070
    PopulateComboBox cboCapValue(cccVolumeUnits), True, "mL|uL|nL|pL", 2   '7080
    
    ' Note: Subs ResizeCapillaryCalcsForm and ShowHideWeightSource use
    '         cConcentrationUnitsFirstWeightIndex to determine whether
    '         weight-based (mg/mL) or mole-based (Molar) mass units are being used
    '       Be sure to change cConcentrationUnitsFirstWeightIndex if necessary if
    '         modifying the following units
    PopulateComboBox cboCapValue(cccMassRateConcentrationUnits), True, "Molar|milliMolar|microMolar|nanoMolar|picoMolar|femtoMolar|attoMolar|mg/mL|ug/mL|ng/mL|ug/uL|ng/uL", 2   '7090
    
    With cboCapValue(cccMassRateVolFlowRateUnits)       ' 7050 Dup
        For intIndex = 0 To cboCapValue(cccFlowRateUnits).ListCount - 1
            .AddItem cboCapValue(cccFlowRateUnits).List(intIndex)
        Next intIndex
        .ListIndex = cboCapValue(cccFlowRateUnits).ListIndex
    End With
    
    With cboCapValue(cccMassRateInjectionTimeUnits)   '7070 Dup
        For intIndex = 0 To cboCapValue(cccDeadTimeUnits).ListCount - 1
            .AddItem cboCapValue(cccDeadTimeUnits).List(intIndex)
        Next intIndex
        .ListIndex = cboCapValue(cccDeadTimeUnits).ListIndex
    End With
    
    PopulateComboBox cboCapValue(cccMassFlowRateUnits), True, "pmol/min|fmol/min|amol/min|pmol/sec|fmol/sec|amol/sec", 4   '7100
    PopulateComboBox cboCapValue(cccMassRateMolesInjectedUnits), True, "Moles|milliMoles|microMoles|nanoMoles|picoMoles|femtoMoles|attoMoles", 5   '7110
    
    With cboCapValue(cccBdLinearVelocityUnits)    '7060 Dup
        For intIndex = 0 To cboCapValue(cccLinearVelocityUnits).ListCount - 1
            .AddItem cboCapValue(cccLinearVelocityUnits).List(intIndex)
        Next intIndex
        .ListIndex = cboCapValue(cccLinearVelocityUnits).ListIndex
    End With

    With cboCapValue(cccBdOpenTubeLengthUnits)    '7035 Dup
        For intIndex = 0 To cboCapValue(cccColumnLengthUnits).ListCount - 1
            .AddItem cboCapValue(cccColumnLengthUnits).List(intIndex)
        Next intIndex
        .ListIndex = cboCapValue(cccColumnLengthUnits).ListIndex
    End With

    With cboCapValue(cccBdOpenTubeIDUnits)    '7035 Dup
        For intIndex = 0 To cboCapValue(cccColumnIDUnits).ListCount - 1
            .AddItem cboCapValue(cccColumnIDUnits).List(intIndex)
        Next intIndex
        .ListIndex = cboCapValue(cccColumnIDUnits).ListIndex
    End With

    With cboCapValue(cccBdInitialPeakWidthUnits)         '7070 Dup
        For intIndex = 0 To cboCapValue(cccDeadTimeUnits).ListCount - 1
            .AddItem cboCapValue(cccDeadTimeUnits).List(intIndex)
        Next intIndex
        .ListIndex = 2
    End With

    With cboCapValue(cccBdResultantPeakWidthUnits)         '7070 Dup
        For intIndex = 0 To cboCapValue(cccBdInitialPeakWidthUnits).ListCount - 1
            .AddItem cboCapValue(cccBdInitialPeakWidthUnits).List(intIndex)
        Next intIndex
        .ListIndex = cboCapValue(cccBdInitialPeakWidthUnits).ListIndex
    End With

End Sub

' Purpose: Position controls on the form
Private Sub PositionFormControls()
    Dim RowSpacing As Integer, LabelAdjust As Integer
    RowSpacing = 360
    LabelAdjust = 30
    
    mDefaultCapValuesLoaded = False

    fraBroadening.Visible = False
    fraWeightSource.Visible = False
  
    cboCapillaryType.Top = 160
    cboCapillaryType.Left = 240
    cboComputationType.Top = 520
    cboComputationType.Left = cboCapillaryType.Left
    
    cmdShowPeakBroadening.Top = cboCapillaryType.Top - 40
    cmdShowPeakBroadening.Left = 5160

    cmdViewEquations.Top = cmdViewEquations.Top
    cmdViewEquations.Left = 7440
    
    cmdOK.Top = 240
    cmdOK.Left = 9360
    
    lblPressure.Top = 1080
    lblPressure.Left = cboCapillaryType.Left
    txtCapValue(cctPressure).Top = lblPressure.Top - LabelAdjust
    txtCapValue(cctPressure).Left = 2640
    cboCapValue(cccPressureUnits).Top = lblPressure.Top - LabelAdjust
    cboCapValue(cccPressureUnits).Left = 3840
    
    lblLength.Top = lblPressure.Top + RowSpacing * 1
    lblLength.Left = cboCapillaryType.Left
    txtCapValue(cctColumnLength).Top = lblLength.Top - LabelAdjust
    txtCapValue(cctColumnLength).Left = txtCapValue(cctPressure).Left
    cboCapValue(cccColumnLengthUnits).Top = lblLength.Top - LabelAdjust
    cboCapValue(cccColumnLengthUnits).Left = cboCapValue(cccPressureUnits).Left
    
    lblColumnID.Top = lblPressure.Top + RowSpacing * 2
    lblColumnID.Left = cboCapillaryType.Left
    txtCapValue(cctColumnID).Top = lblColumnID.Top - LabelAdjust
    txtCapValue(cctColumnID).Left = txtCapValue(cctPressure).Left
    cboCapValue(cccColumnIDUnits).Top = lblColumnID.Top - LabelAdjust
    cboCapValue(cccColumnIDUnits).Left = cboCapValue(cccPressureUnits).Left
    
    lblViscosity.Top = lblPressure.Top + RowSpacing * 3
    lblViscosity.Left = cboCapillaryType.Left
    txtCapValue(cctViscosity).Top = lblViscosity.Top - LabelAdjust
    txtCapValue(cctViscosity).Left = txtCapValue(cctPressure).Left
    cboCapValue(cccViscosityUnits).Top = lblViscosity.Top - LabelAdjust
    cboCapValue(cccViscosityUnits).Left = cboCapValue(cccPressureUnits).Left
    
    lblParticleDiameter.Top = lblPressure.Top + RowSpacing * 4
    lblParticleDiameter.Left = cboCapillaryType.Left
    txtCapValue(cctParticleDiamter).Top = lblParticleDiameter.Top - LabelAdjust
    txtCapValue(cctParticleDiamter).Left = txtCapValue(cctPressure).Left
    cboCapValue(cccParticleDiameterUnits).Top = lblParticleDiameter.Top - LabelAdjust
    cboCapValue(cccParticleDiameterUnits).Left = cboCapValue(cccPressureUnits).Left
    
        lblFlowRateLabel.Top = lblPressure.Top
        lblFlowRateLabel.Left = 5880
        txtCapValue(cctFlowRate).Top = lblFlowRateLabel.Top - LabelAdjust
        txtCapValue(cctFlowRate).Left = 8280
        cboCapValue(cccFlowRateUnits).Top = lblFlowRateLabel.Top - LabelAdjust
        cboCapValue(cccFlowRateUnits).Left = 9480
        
        lblLinearVelocityLabel.Top = lblFlowRateLabel.Top + RowSpacing
        lblLinearVelocityLabel.Left = lblFlowRateLabel.Left
        lblLinearVelocity.Top = lblLinearVelocityLabel.Top
        lblLinearVelocity.Left = txtCapValue(cctFlowRate).Left + 40
        cboCapValue(cccLinearVelocityUnits).Top = lblLinearVelocityLabel.Top - LabelAdjust
        cboCapValue(cccLinearVelocityUnits).Left = cboCapValue(cccFlowRateUnits).Left
        
        lblDeadTimeLabel.Top = lblFlowRateLabel.Top + RowSpacing * 2
        lblDeadTimeLabel.Left = lblFlowRateLabel.Left
        txtCapValue(cctDeadTime).Top = lblDeadTimeLabel.Top - LabelAdjust
        txtCapValue(cctDeadTime).Left = txtCapValue(cctFlowRate).Left
        cboCapValue(cccDeadTimeUnits).Top = lblDeadTimeLabel.Top - LabelAdjust
        cboCapValue(cccDeadTimeUnits).Left = cboCapValue(cccFlowRateUnits).Left
    
        lblVolumeLabel.Top = lblFlowRateLabel.Top + RowSpacing * 3
        lblVolumeLabel.Left = lblFlowRateLabel.Left
        lblVolume.Top = lblVolumeLabel.Top
        lblVolume.Left = txtCapValue(cctFlowRate).Left + 40
        cboCapValue(cccVolumeUnits).Top = lblVolumeLabel.Top - LabelAdjust
        cboCapValue(cccVolumeUnits).Left = cboCapValue(cccFlowRateUnits).Left
    
        lblPorosity.Top = lblFlowRateLabel.Top + RowSpacing * 4
        lblPorosity.Left = lblFlowRateLabel.Left
        txtCapValue(cctPorosity).Top = lblPorosity.Top - LabelAdjust
        txtCapValue(cctPorosity).Left = cboCapValue(cccFlowRateUnits).Left
    
    cmdComputeViscosity.Top = lblPressure.Top + RowSpacing * 5
    cmdComputeViscosity.Left = 240
    
    fraMassRate.Top = cmdComputeViscosity.Top + cmdComputeViscosity.Height + 140
    fraMassRate.Left = 120
    
    lblMassRateConcentration.Top = 300
    lblMassRateConcentration.Left = 240
    txtCapValue(cctMassRateConcentration).Top = lblMassRateConcentration.Top - LabelAdjust
    txtCapValue(cctMassRateConcentration).Left = 2520
    cboCapValue(cccMassRateConcentrationUnits).Top = lblMassRateConcentration.Top - LabelAdjust
    cboCapValue(cccMassRateConcentrationUnits).Left = 3720
    
    lblMassRateVolFlowRate.Top = lblMassRateConcentration.Top + RowSpacing
    lblMassRateVolFlowRate.Left = lblMassRateConcentration.Left
    txtCapValue(cctMassRateVolFlowRate).Top = lblMassRateVolFlowRate.Top + 80
    txtCapValue(cctMassRateVolFlowRate).Left = txtCapValue(cctMassRateConcentration).Left
    cboCapValue(cccMassRateVolFlowRateUnits).Top = txtCapValue(cctMassRateVolFlowRate).Top
    cboCapValue(cccMassRateVolFlowRateUnits).Left = cboCapValue(cccMassRateConcentrationUnits).Left
    chkMassRateLinkFlowRate.Top = lblMassRateVolFlowRate.Top + 240
    chkMassRateLinkFlowRate.Left = lblMassRateConcentration.Left + 240
    
    lblMassRateInjectionTime.Top = lblMassRateVolFlowRate.Top + 600
    lblMassRateInjectionTime.Left = lblMassRateConcentration.Left
    txtCapValue(cctMassRateInjectionTime).Top = lblMassRateInjectionTime.Top - LabelAdjust
    txtCapValue(cctMassRateInjectionTime).Left = txtCapValue(cctMassRateConcentration).Left
    cboCapValue(cccMassRateInjectionTimeUnits).Top = lblMassRateInjectionTime.Top - LabelAdjust
    cboCapValue(cccMassRateInjectionTimeUnits).Left = cboCapValue(cccMassRateConcentrationUnits).Left
    
        lblMassFlowRateLabel.Top = lblMassRateConcentration.Top
        lblMassFlowRateLabel.Left = 5760
        lblMassFlowRate.Top = lblMassFlowRateLabel.Top
        lblMassFlowRate.Left = 8040
        cboCapValue(cccMassFlowRateUnits).Top = lblMassFlowRateLabel.Top - LabelAdjust
        cboCapValue(cccMassFlowRateUnits).Left = 9200
        
        lblMolesInjectedLabel.Top = lblMassRateInjectionTime.Top
        lblMolesInjectedLabel.Left = lblMassFlowRateLabel.Left
        lblMolesInjected.Top = lblMolesInjectedLabel.Top
        lblMolesInjected.Left = lblMassFlowRate.Left
        cboCapValue(cccMassRateMolesInjectedUnits).Top = lblMolesInjectedLabel.Top - LabelAdjust
        cboCapValue(cccMassRateMolesInjectedUnits).Left = cboCapValue(cccMassFlowRateUnits).Left
    
    fraBroadening.Top = fraMassRate.Top
    fraBroadening.Left = fraMassRate.Left
    
    lblBdLinearVelocity.Top = lblMassRateConcentration.Top
    lblBdLinearVelocity.Left = lblMassRateConcentration.Left
    txtCapValue(cctBdLinearVelocity).Top = lblBdLinearVelocity.Top + 80
    txtCapValue(cctBdLinearVelocity).Left = txtCapValue(cctMassRateConcentration).Left
    cboCapValue(cccBdLinearVelocityUnits).Top = txtCapValue(cctBdLinearVelocity).Top
    cboCapValue(cccBdLinearVelocityUnits).Left = cboCapValue(cccMassRateConcentrationUnits).Left
    chkBdLinkLinearVelocity.Top = lblBdLinearVelocity.Top + 240
    chkBdLinkLinearVelocity.Left = lblBdLinearVelocity.Left + 240

    lblDiffusionCoefficient.Top = lblBdLinearVelocity.Top + 600
    lblDiffusionCoefficient.Left = lblBdLinearVelocity.Left
    txtCapValue(cctBdDiffusionCoefficient).Top = lblDiffusionCoefficient.Top - LabelAdjust
    txtCapValue(cctBdDiffusionCoefficient).Left = txtCapValue(cctBdLinearVelocity).Left
    rtfDiffusionCoefficient.Top = lblDiffusionCoefficient.Top - LabelAdjust
    rtfDiffusionCoefficient.Left = cboCapValue(cccBdLinearVelocityUnits).Left

    lblBdOpenTubeLength.Top = lblDiffusionCoefficient.Top + RowSpacing
    lblBdOpenTubeLength.Left = lblBdLinearVelocity.Left
    txtCapValue(cctBdOpenTubeLength).Top = lblBdOpenTubeLength.Top - LabelAdjust
    txtCapValue(cctBdOpenTubeLength).Left = txtCapValue(cctBdLinearVelocity).Left
    cboCapValue(cccBdOpenTubeLengthUnits).Top = lblBdOpenTubeLength.Top - LabelAdjust
    cboCapValue(cccBdOpenTubeLengthUnits).Left = cboCapValue(cccBdLinearVelocityUnits).Left
    
    lblBdOpenTubeID.Top = lblBdOpenTubeLength.Top + RowSpacing
    lblBdOpenTubeID.Left = lblBdLinearVelocity.Left
    txtCapValue(cctBdOpenTubeID).Top = lblBdOpenTubeID.Top - LabelAdjust
    txtCapValue(cctBdOpenTubeID).Left = txtCapValue(cctBdLinearVelocity).Left
    cboCapValue(cccBdOpenTubeIDUnits).Top = lblBdOpenTubeID.Top - LabelAdjust
    cboCapValue(cccBdOpenTubeIDUnits).Left = cboCapValue(cccBdLinearVelocityUnits).Left

    lblBdInitialPeakWidth.Top = lblBdOpenTubeLength.Top + RowSpacing * 2
    lblBdInitialPeakWidth.Left = lblBdLinearVelocity.Left
    txtCapValue(cctBdInitialPeakWidth).Top = lblBdInitialPeakWidth.Top - LabelAdjust
    txtCapValue(cctBdInitialPeakWidth).Left = txtCapValue(cctBdLinearVelocity).Left
    cboCapValue(cccBdInitialPeakWidthUnits).Top = lblBdInitialPeakWidth.Top - LabelAdjust
    cboCapValue(cccBdInitialPeakWidthUnits).Left = cboCapValue(cccBdLinearVelocityUnits).Left

        lblOptimumLinearVelocityLabel.Top = lblBdLinearVelocity.Top
        lblOptimumLinearVelocityLabel.Left = lblMassFlowRateLabel.Left
        lblOptimumLinearVelocity.Top = lblOptimumLinearVelocityLabel.Top
        lblOptimumLinearVelocity.Left = 8160
        lblOptimumLinearVelocityUnit.Top = lblOptimumLinearVelocityLabel.Top
        lblOptimumLinearVelocityUnit.Left = 9360
        lblOptimumLinearVelocityBasis.Top = chkBdLinkLinearVelocity.Top
        lblOptimumLinearVelocityBasis.Left = lblOptimumLinearVelocityLabel.Left + 100

        lblBdTemporalVarianceLabel.Top = lblDiffusionCoefficient.Top
        lblBdTemporalVarianceLabel.Left = lblOptimumLinearVelocityLabel.Left
        lblBdTemporalVariance.Top = lblBdTemporalVarianceLabel.Top
        lblBdTemporalVariance.Left = lblOptimumLinearVelocity.Left
        rtfBdTemporalVarianceUnit.Top = lblBdTemporalVarianceLabel.Top - LabelAdjust
        rtfBdTemporalVarianceUnit.Left = lblOptimumLinearVelocityUnit.Left

        lblBdAdditionalVarianceLabel.Top = lblBdTemporalVarianceLabel.Top + RowSpacing
        lblBdAdditionalVarianceLabel.Left = lblBdTemporalVarianceLabel.Left
        txtCapValue(cctBdAdditionalVariance).Top = lblBdAdditionalVarianceLabel.Top - LabelAdjust
        txtCapValue(cctBdAdditionalVariance).Left = lblBdTemporalVariance.Left
        rtfBdAdditionalVarianceUnit.Top = txtCapValue(cctBdAdditionalVariance).Top
        rtfBdAdditionalVarianceUnit.Left = rtfBdTemporalVarianceUnit.Left

        lblBdResultantPeakWidthLabel.Top = lblBdOpenTubeID.Top
        lblBdResultantPeakWidthLabel.Left = lblBdTemporalVarianceLabel.Left
        lblBdResultantPeakWidth.Top = lblBdResultantPeakWidthLabel.Top
        lblBdResultantPeakWidth.Left = lblBdTemporalVariance.Left
        cboCapValue(cccBdResultantPeakWidthUnits).Top = lblBdResultantPeakWidthLabel.Top - LabelAdjust
        cboCapValue(cccBdResultantPeakWidthUnits).Left = rtfBdTemporalVarianceUnit.Left - 120

        lblBdPercentVarianceIncreaseLabel.Top = lblBdInitialPeakWidth.Top
        lblBdPercentVarianceIncreaseLabel.Left = lblBdTemporalVarianceLabel.Left
        lblBdPercentVarianceIncrease.Top = lblBdPercentVarianceIncreaseLabel.Top
        lblBdPercentVarianceIncrease.Left = lblBdTemporalVariance.Left
        
    fraWeightSource.Top = fraMassRate.Top + fraMassRate.Height + 100
    fraWeightSource.Left = 600
    PositionWeightSourceframeControls Me

End Sub

' Purpose: Resize form
Private Sub ResizeCapillaryCalcsForm(boolResizeToDefaultHeight As Boolean)
    Dim lngHeightToSet As Long
    
    If fraBroadening.Visible = True Then
        lngHeightToSet = CAPILLARY_CALCS_FORM_INITIAL_HEIGHT + 1150
    Else
        If cboCapValue(cccMassRateConcentrationUnits).ListIndex >= cConcentrationUnitsFirstWeightIndex Then
            ' Weight-based mass rate units
            lngHeightToSet = CAPILLARY_CALCS_FORM_INITIAL_HEIGHT + 1400
        Else
            ' Mole-based mass rate units
            lngHeightToSet = CAPILLARY_CALCS_FORM_INITIAL_HEIGHT
        End If
    End If
    
    If Me.WindowState = vbNormal Then
        If Me.Width > CAPILLARY_CALCS_FORM_INITIAL_WIDTH Then
            Me.Width = CAPILLARY_CALCS_FORM_INITIAL_WIDTH
        End If
    
        If boolResizeToDefaultHeight Then
            Me.Height = lngHeightToSet
        Else
            If Me.Height > lngHeightToSet Then
                Me.Height = lngHeightToSet
            End If
        End If
    End If
    
End Sub

' Purpose: Saved changed combo values to OpenCapValues or PackedCapValues
Private Sub SaveChangedComboValue(cboThisComboBox As ComboBox, Index As Integer)
    
    If Not mDefaultCapValuesLoaded Then Exit Sub
    
    If cboThisComboBox.Name = cboComputationType.Name Then
        gCapFlowComputationTypeSave = cboThisComboBox.ListIndex
    Else
        If cboCapillaryType.ListIndex = ctOpenTubularCapillary Then
            ' Open Capillary
            OpenCapVals.ComboValues(Index) = cboThisComboBox.ListIndex
        Else
            ' Packed Capillary
            PackedCapVals.ComboValues(Index) = cboThisComboBox.ListIndex
        End If
    End If
            
End Sub

' Purpose: Save the value in OpenCapVals or PackedCapVals to allow for easy switching between open and packed capillaries
Private Sub SaveChangedTextValue(Index As Integer)
    
    If mUpdatingCapValues Or Not mDefaultCapValuesLoaded Then Exit Sub
    
    On Error Resume Next
    
    If cboCapillaryType.ListIndex = ctOpenTubularCapillary Then
        ' Open Capillary
        OpenCapVals.TextValues(Index) = CDblSafe(txtCapValue(Index).Text)
    Else
        ' Packed Capillary
        PackedCapVals.TextValues(Index) = CDblSafe(txtCapValue(Index).Text)
    End If
    
End Sub

Private Sub ShowHideBroadeningFrame()
    With cmdShowPeakBroadening
        If fraBroadening.Visible = False Then
            fraBroadening.Visible = True
            gCapFlowShowPeakBroadeningSave = 1
        Else
            fraBroadening.Visible = False
            gCapFlowShowPeakBroadeningSave = 0
        End If
    End With
    
    ShowHideWeightSource
    ResizeCapillaryCalcsForm True

End Sub

' Purpose: Displays fraWeightSource if the concentration units involve mass (e.g. g
Private Sub ShowHideWeightSource()
    If cboCapValue(cccMassRateConcentrationUnits).ListIndex >= cConcentrationUnitsFirstWeightIndex And _
        fraBroadening.Visible = False Then
        fraWeightSource.Visible = True
    Else
        fraWeightSource.Visible = False
    End If
    ResizeCapillaryCalcsForm True
    
End Sub

Private Sub SynchronizeValueWithDll(eTextBoxID As cctCapCalcTextBoxIDConstants)
    Dim dblNewValue As Double

    If CheckDoNotModify(eTextBoxID) Then Exit Sub

    ' Update value in objCapillaryFlow
    dblNewValue = CDblSafe(txtCapValue(eTextBoxID))
    
    Select Case eTextBoxID
    Case cctPressure: objCapillaryFlow.SetBackPressure dblNewValue, cboCapValue(cccPressureUnits).ListIndex
    Case cctColumnLength: objCapillaryFlow.SetColumnLength dblNewValue, cboCapValue(cccColumnLengthUnits).ListIndex
    Case cctColumnID: objCapillaryFlow.SetColumnID dblNewValue, cboCapValue(cccColumnIDUnits).ListIndex
    Case cctViscosity: objCapillaryFlow.SetSolventViscosity dblNewValue, cboCapValue(cccViscosityUnits).ListIndex
    Case cctParticleDiamter: objCapillaryFlow.SetParticleDiameter dblNewValue, cboCapValue(cccParticleDiameterUnits).ListIndex
    Case cctFlowRate: objCapillaryFlow.SetVolFlowRate dblNewValue, cboCapValue(cccFlowRateUnits).ListIndex
    Case cctDeadTime: objCapillaryFlow.SetDeadTime dblNewValue, cboCapValue(cccDeadTimeUnits).ListIndex
    Case cctPorosity: objCapillaryFlow.SetInterparticlePorosity dblNewValue
    Case cctMassRateConcentration: objCapillaryFlow.SetMassRateConcentration dblNewValue, cboCapValue(cccMassRateConcentrationUnits).ListIndex
    Case cctMassRateVolFlowRate: objCapillaryFlow.SetMassRateVolFlowRate dblNewValue, cboCapValue(cccMassRateVolFlowRateUnits).ListIndex
    Case cctMassRateInjectionTime: objCapillaryFlow.SetMassRateInjectionTime dblNewValue, cboCapValue(cccMassRateInjectionTimeUnits).ListIndex
    Case cctBdLinearVelocity: objCapillaryFlow.SetExtraColumnBroadeningLinearVelocity dblNewValue, cboCapValue(cccBdLinearVelocityUnits).ListIndex
    Case cctBdDiffusionCoefficient: objCapillaryFlow.SetExtraColumnBroadeningDiffusionCoefficient dblNewValue, udcCmSquaredPerSec
    Case cctBdOpenTubeLength: objCapillaryFlow.SetExtraColumnBroadeningOpenTubeLength dblNewValue, cboCapValue(cccBdOpenTubeLengthUnits).ListIndex
    Case cctBdOpenTubeID: objCapillaryFlow.SetExtraColumnBroadeningOpenTubeID dblNewValue, cboCapValue(cccBdOpenTubeIDUnits).ListIndex
    Case cctBdInitialPeakWidth: objCapillaryFlow.SetExtraColumnBroadeningInitialPeakWidthAtBase dblNewValue, cboCapValue(cccBdInitialPeakWidthUnits).ListIndex
    Case cctBdAdditionalVariance: objCapillaryFlow.SetExtraColumnBroadeningAdditionalVariance dblNewValue
    Case Else
        ' This shouldn't happen
        Debug.Assert False
    End Select

End Sub

' Switch capillary types
Public Sub UpdateCapillaryType()
    
    Dim blnShowControls As Boolean
    
    On Error GoTo CCErrorHandler
        
    Select Case cboCapillaryType.ListIndex
    Case ctOpenTubularCapillary
        ' Set Packed Capillary Fields to be invisible
        blnShowControls = False
        objCapillaryFlow.SetCapillaryType ctOpenTubularCapillary
    Case Else
        ' Set Packed Capillary Fields to be visible
        blnShowControls = True
        objCapillaryFlow.SetCapillaryType ctPackedCapillary
    End Select
    
    lblParticleDiameter.Visible = blnShowControls
    txtCapValue(cctParticleDiamter).Visible = blnShowControls
    cboCapValue(cccParticleDiameterUnits).Visible = blnShowControls
    
    lblPorosity.Visible = blnShowControls
    txtCapValue(cctPorosity).Visible = blnShowControls
    
    LoadCapFormValues
    CopyFlowRate
    FindDesiredValue
    
CCStart:
    'This code skips the error handler
    Exit Sub

CCErrorHandler:
    GeneralErrorHandler "CapillaryCalcs|cboCapillarType_Click", Err.Number, Err.Description
    Resume CCStart

End Sub

Private Sub cboCapillaryType_Click()
    UpdateCapillaryType
End Sub

Private Sub cboCapValue_Click(Index As Integer)
    Dim eCurrentUnits As ulnUnitsLengthConstants
    Dim eNewUnits As ulnUnitsLengthConstants
    Dim dblNewColumnID As Double
    
    Select Case Index
    Case cccPressureUnits:              SynchronizeValueWithDll cctPressure
    Case cccColumnLengthUnits:          SynchronizeValueWithDll cctColumnLength
    Case cccColumnIDUnits:              SynchronizeValueWithDll cctColumnID
    Case cccViscosityUnits:             SynchronizeValueWithDll cctViscosity
    Case cccParticleDiameterUnits:      SynchronizeValueWithDll cctParticleDiamter
    Case cccFlowRateUnits:              SynchronizeValueWithDll cctFlowRate
    Case cccLinearVelocityUnits                 ' Nothing needs to be updated in the Dll
    Case cccDeadTimeUnits:              SynchronizeValueWithDll cctDeadTime
    Case cccVolumeUnits                         ' Nothing needs to be updated in the Dll
    Case cccMassRateConcentrationUnits: SynchronizeValueWithDll cctMassRateConcentration
    Case cccMassRateVolFlowRateUnits:   SynchronizeValueWithDll cctMassRateVolFlowRate
    Case cccMassRateInjectionTimeUnits: SynchronizeValueWithDll cctMassRateInjectionTime
    Case cccMassFlowRateUnits                   ' Nothing needs to be updated in the Dll
    Case cccMassRateMolesInjectedUnits          ' Nothing needs to be updated in the Dll
    Case cccBdLinearVelocityUnits:      SynchronizeValueWithDll cctBdLinearVelocity
    Case cccBdOpenTubeLengthUnits:      SynchronizeValueWithDll cctBdOpenTubeLength
    Case cccBdOpenTubeIDUnits:          SynchronizeValueWithDll cctBdOpenTubeID
    Case cccBdInitialPeakWidthUnits:    SynchronizeValueWithDll cctBdInitialPeakWidth
    Case cccBdResultantPeakWidthUnits           ' Nothing needs to be updated in the Dll
    Case Else
        ' This shouldn't happen
        Debug.Assert False
    End Select
    
    If Index = cccColumnIDUnits And Not mUpdatingCapValues Then
        ' Change the units from um to inches or back (if required)
        If cboCapValue(Index).ListIndex <> eColumnIDUnitsIndexSaved And mDefaultCapValuesLoaded Then
            eCurrentUnits = eColumnIDUnitsIndexSaved
            eNewUnits = cboCapValue(Index).ListIndex
            
            dblNewColumnID = objCapillaryFlow.ConvertLength(txtCapValue(cctColumnID).Text, eCurrentUnits, eNewUnits)
            If eNewUnits = ulnInches Then
                txtCapValue(cctColumnID).Text = Format(dblNewColumnID, "0.000000")
            Else
                If dblNewColumnID < 0.01 Then
                    txtCapValue(cctColumnID).Text = Format(dblNewColumnID, "0.000E+00")
                Else
                    txtCapValue(cctColumnID).Text = Format(dblNewColumnID, "0.000")
                End If
            End If
        End If
        eColumnIDUnitsIndexSaved = cboCapValue(Index).ListIndex
    End If
    
    If Index <= cccVolumeUnits Then
        If Index = cccFlowRateUnits And cChkBox(chkMassRateLinkFlowRate) Then
            CopyFlowRate
            FindDesiredValue
            FindMassRate
            FindBroadening
        Else
            FindDesiredValue
        End If
    ElseIf Index <= cccMassRateMolesInjectedUnits Then
        FindMassRate
    Else
        If Index = cccBdLinearVelocityUnits Then
            If cChkBox(chkBdLinkLinearVelocity) Then CopyLinearVelocity
        End If
        FindBroadening
    End If
    
    If Index = cccLinearVelocityUnits Then
        FindOptimumLinearVelocity
    End If
    
    If Index = cccMassRateConcentrationUnits Then
        ShowHideWeightSource
    End If

    SaveChangedComboValue cboCapValue(Index), Index
        
End Sub

Private Sub cboCapValue_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim Cancel As Boolean
    cboCapValue_Validate Index, Cancel

End Sub

Private Sub cboCapValue_Validate(Index As Integer, Cancel As Boolean)
    If Index = cccMassRateVolFlowRateUnits Then
        CopyFlowRate
        FindMassRate
    ElseIf Index = cccBdLinearVelocityUnits Then
        If cChkBox(chkBdLinkLinearVelocity) Then CopyLinearVelocity
        FindBroadening
    End If
End Sub

Private Sub cboComputationType_Click()
    HighlightTargetBoxes
    CopyFlowRate
    FindDesiredValue
    SaveChangedComboValue cboComputationType, 0
End Sub

Private Sub chkBdLinkLinearVelocity_Click()
    If cChkBox(chkBdLinkLinearVelocity) Then CopyLinearVelocity
    FindBroadening
    gCapFlowLinkBdLinearVelocitySave = chkBdLinkLinearVelocity.value
    
End Sub

Private Sub chkMassRateLinkFlowRate_Click()
    CopyFlowRate
    FindMassRate
    gCapFlowLinkMassRateFlowRateSave = chkMassRateLinkFlowRate.value
End Sub

Private Sub cmdComputeViscosity_Click()
    frmViscosityForMeCN.Show
End Sub

Private Sub cmdOK_Click()
    HideFormShowMain Me
End Sub

Private Sub cmdShowPeakBroadening_Click()
    ShowHideBroadeningFrame
End Sub

Private Sub cmdViewBroadeningEquations_Click()
    frmEquationsBroadening.Show vbModal
End Sub

Private Sub cmdViewEquations_Click()
    If cboCapillaryType.ListIndex = ctPackedCapillary Then
        ' Packed Capillary
        frmEquationsPackedCapillary.Show vbModal
    Else
        ' Open Capillary
        frmEquationsOpenTube.Show vbModal
    End If
    
End Sub

Private Sub Form_Activate()
    SizeAndCenterWindow Me, cWindowTopCenter, CAPILLARY_CALCS_FORM_INITIAL_WIDTH, CAPILLARY_CALCS_FORM_INITIAL_HEIGHT
    
    DisplayCurrentFormulaOnSubForm Me
    
    UpdateCapillaryType
    CopyFlowRate
    FindDesiredValue
    FindMassRate
    FindBroadening
    
    PossiblyHideMainWindow
    
    FillRTFBoxes
End Sub

Private Sub Form_Load()
    
    ' Turn off auto-compute
    objCapillaryFlow.SetAutoComputeEnabled False
    
    PositionFormControls
    
    PopulateComboBoxes
        
    ' Make sure appropriate Weight Source box is shown
    '  Note: also calls ResizeCapillaryCalcsForm
    ShowHideWeightSource
    
    ' Also make sure correct mass input controls are shown
    ShowHideMassInputControlsGlobal Me
    
    mDefaultCapValuesLoaded = True
    UpdateCapillaryType

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    QueryUnloadFormHandler Me, Cancel, UnloadMode
End Sub

Private Sub Form_Resize()
    ResizeCapillaryCalcsForm False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objCapillaryFlow = Nothing
End Sub

Private Sub lblLinearVelocity_Change()
    If cChkBox(chkBdLinkLinearVelocity) Then
        CopyLinearVelocity
        FindBroadening
    End If
End Sub

Private Sub lblMWT_Change()
    FindDesiredValue
End Sub

Private Sub mnuCapillaryFlowHelp_Click()
    ShowHelpPage hwnd, 3070
End Sub

Private Sub mnuClose_Click()
    HideFormShowMain Me
End Sub

Private Sub mnuLoadCapValues_Click()
    LoadCapillaryFlowInfo
    
    ' Update controls to correct values
    LoadCapFormValues
    FindDesiredValue

End Sub

Private Sub mnuSaveCapValues_Click()
    SaveCapillaryFlowInfo

End Sub

Private Sub optWeightSource_Click(Index As Integer)
    ShowHideMassInputControlsGlobal Me
    FindMassRate
End Sub

Private Sub txtCapValue_Change(Index As Integer)
    
    SaveChangedTextValue Index
    If CheckDoNotModify(CInt(Index)) Then Exit Sub
    
    SynchronizeValueWithDll CInt(Index)
    
    If Index <= cctPorosity Then
        CopyFlowRate
        FindDesiredValue
        If cboComputationType.ListIndex = caFindVolFlowRateUsingPressure Or cboComputationType.ListIndex = caFindVolFlowUsingDeadTime Then
            CopyFlowRate
        End If
    ElseIf Index <= cctMassRateInjectionTime Then
        If Index = cctMassRateVolFlowRate And cChkBox(chkMassRateLinkFlowRate) Then
            CopyFlowRate
            FindDesiredValue
        End If
        FindMassRate
    Else
        If Index = cctBdLinearVelocity And cChkBox(chkBdLinkLinearVelocity) Then
            CopyLinearVelocity
            FindDesiredValue
        End If
        FindBroadening
    End If
    
    If Index = cctParticleDiamter Or Index = cctBdDiffusionCoefficient Then
        FindOptimumLinearVelocity
    End If
    
End Sub

Private Sub txtCapValue_GotFocus(Index As Integer)
    HighlightOnFocus txtCapValue(Index)
    
End Sub

Private Sub txtCapValue_KeyPress(Index As Integer, KeyAscii As Integer)
            
    If CheckDoNotModify(CInt(Index)) Then
        ' Check for Ctrl+C (copy) and Ctrl+A (select all)
        If KeyAscii <> 3 And KeyAscii <> 1 Then
            KeyAscii = 0
            Exit Sub
        End If
    End If
    
    TextBoxKeyPressHandler txtCapValue(Index), KeyAscii, True, True, True, False, True, False, False, False, False, True
End Sub

Private Sub txtCustomMass_Change()
    FindMassRate
End Sub

Private Sub txtCustomMass_GotFocus()
    HighlightOnFocus txtCustomMass

End Sub

Private Sub txtCustomMass_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtCustomMass, KeyAscii, True, True, True, False, True, False, False, False, False, True
End Sub
