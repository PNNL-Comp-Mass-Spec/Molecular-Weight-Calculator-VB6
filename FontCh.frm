VERSION 5.00
Begin VB.Form frmChangeFont 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change Formula Font"
   ClientHeight    =   1440
   ClientLeft      =   3090
   ClientTop       =   1305
   ClientWidth     =   5025
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
   HelpContextID   =   4060
   Icon            =   "FONTCH.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1440
   ScaleWidth      =   5025
   Tag             =   "8000"
   Begin VB.ComboBox cboFontSize 
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
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Tag             =   "8020"
      Top             =   960
      Width           =   1335
   End
   Begin VB.ComboBox cboFonts 
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
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Tag             =   "8010"
      Top             =   480
      Width           =   4575
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "Cl&ose"
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
      Height          =   360
      Left            =   3600
      TabIndex        =   2
      Tag             =   "4000"
      Top             =   960
      Width           =   1155
   End
   Begin VB.Label lblFontSize 
      Caption         =   "Font Size:"
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
      TabIndex        =   4
      Tag             =   "8070"
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label lblFonts 
      Caption         =   "Change Formula Font to:"
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
      TabIndex        =   3
      Tag             =   "8060"
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "frmChangeFont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MIN_FONT_SIZE = 6

Private Sub LoadFontsIntoComboBox()
    Dim intIndex As Integer, intFontCount As Integer
    
    intFontCount = Screen.FontCount - 1
    
    'If intFontCount > 25 Then frmMain.lblStatus.Caption = "Building font list 0%"
    ' Load Fonts in the Combo Box
    For intIndex = 0 To intFontCount
        cboFonts.AddItem Screen.Fonts(intIndex)
        If intIndex Mod 25 = 0 Then
            frmMain.lblStatus.Caption = "Building font list " & _
                                        Trim(Str(CIntSafeDbl(intIndex / intFontCount * 100))) & "%"
            DoEvents
        End If
    Next intIndex
    
    With cboFontSize
        For intIndex = MIN_FONT_SIZE To 14
            .AddItem intIndex
            If intIndex = objMwtWin.RtfFontSize Then .ListIndex = .ListCount - 1
        Next intIndex
    
''  ' In order to use larger fonts, I need to make rtfFormula().Height be dependent on a variable
''  ' This variable would also have to be used by the frmMain.PositionFormControls() function
''        For intIndex = 16 To 32 Step 2
''            .AddItem intIndex
''            If intIndex = objMwtWin.RtfFontSize Then .ListIndex = .ListCount - 1
''        Next intIndex
    End With

End Sub

Private Sub PositionFormControls()
    Me.Caption = LookupLanguageCaption(8000, "Change Formula Font")
    CmdOK.Caption = LookupLanguageCaption(4000, "Cl&ose")
    
    lblFonts.Top = 120
    lblFonts.Left = 240
    lblFonts.Caption = LookupLanguageCaption(8060, "Change Formula Font to:")
    cboFonts.Top = 480
    cboFonts.Left = lblFonts.Left

    lblFontSize.Top = 960
    lblFontSize.Left = lblFonts.Left
    lblFontSize.Caption = LookupLanguageCaption(8070, "Font Size:")
    
    cboFontSize.Top = lblFontSize.Top
    cboFontSize.Left = (frmChangeFont.ScaleWidth - CmdOK.Width) / 4
    cboFontSize.Left = (frmChangeFont.ScaleWidth - CmdOK.Width) / 4
    
    CmdOK.Top = lblFontSize.Top
    CmdOK.Left = (frmChangeFont.ScaleWidth - CmdOK.Width) * 3 / 4

End Sub

Private Sub cboFonts_Click()
    SetFonts cboFonts.Text, objMwtWin.RtfFontSize
End Sub

Private Sub cboFontSize_Click()
    SetFonts cboFonts.Text, Val(cboFontSize.Text)
End Sub

Private Sub cmdOK_Click()
    frmChangeFont.Hide
End Sub

Private Sub Form_Activate()
    ' Put window in center of screen
    SizeAndCenterWindow Me, cWindowExactCenter, 5200, 1850
    
    Dim intIndex As Integer

    ' Display the current font in the combo box
    For intIndex = 0 To Screen.FontCount - 1
        If cboFonts.List(intIndex) = objMwtWin.RtfFontName Then
            cboFonts.ListIndex = intIndex
            Exit For
        End If
    Next intIndex

End Sub

Private Sub Form_Load()
    
    ' Position Form Controls
    PositionFormControls
    
    ' Change mouse pointer to hourglass
    MousePointer = vbHourglass

    ' Load fonts
    LoadFontsIntoComboBox
    
    frmMain.LabelStatus
    
    ' Change mouse pointer back to default
    MousePointer = vbDefault

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    QueryUnloadFormHandler Me, Cancel, UnloadMode
End Sub
