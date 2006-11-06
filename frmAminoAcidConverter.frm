VERSION 5.00
Begin VB.Form frmAminoAcidConverter 
   Caption         =   "Amino Acid Notation Converter"
   ClientHeight    =   4905
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8460
   HelpContextID   =   3055
   Icon            =   "frmAminoAcidConverter.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4905
   ScaleWidth      =   8460
   StartUpPosition =   3  'Windows Default
   Tag             =   "6000"
   Begin VB.ComboBox cboTargetFormula 
      Height          =   315
      Left            =   3960
      Style           =   2  'Dropdown List
      TabIndex        =   10
      ToolTipText     =   "Units of amount to convert from or use for molarity calculation"
      Top             =   4320
      Width           =   735
   End
   Begin VB.CommandButton cmdCopyToFragmentationModeller 
      Caption         =   "Copy to &Fragmentation Modeller"
      Height          =   720
      Left            =   120
      TabIndex        =   8
      Tag             =   "6060"
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CheckBox chkSpaceEvery10 
      Caption         =   "&Add space every 10 residues"
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Tag             =   "6080"
      Top             =   840
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.CheckBox chkSeparateWithDash 
      Caption         =   "&Separate residues with dash"
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Tag             =   "6090"
      Top             =   3000
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.CommandButton cmdCopy3Letter 
      Caption         =   "&Copy 3 letter sequence to formula:"
      Height          =   720
      Left            =   2040
      TabIndex        =   9
      Tag             =   "6050"
      Top             =   4080
      Width           =   1815
   End
   Begin VB.TextBox txt1LetterSequence 
      Height          =   1245
      HelpContextID   =   3055
      Left            =   1920
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Tag             =   "6030"
      Text            =   "frmAminoAcidConverter.frx":08CA
      ToolTipText     =   "Enter sequence using 1 letter abbreviations here"
      Top             =   120
      Width           =   6375
   End
   Begin VB.TextBox txt3LetterSequence 
      Height          =   1845
      HelpContextID   =   3055
      Left            =   1920
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Tag             =   "6040"
      Text            =   "frmAminoAcidConverter.frx":08CC
      ToolTipText     =   "Enter sequence using 3 letter abbreviations here"
      Top             =   2160
      Width           =   6375
   End
   Begin VB.CommandButton cmdConvertTo3Letter 
      Appearance      =   0  'Flat
      Height          =   585
      Left            =   3720
      Picture         =   "frmAminoAcidConverter.frx":08D0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton cmdConvertTo1Letter 
      Height          =   600
      Left            =   5520
      Picture         =   "frmAminoAcidConverter.frx":0D07
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Cl&ose"
      Height          =   480
      Left            =   7080
      TabIndex        =   11
      Tag             =   "4000"
      Top             =   4200
      Width           =   1155
   End
   Begin VB.Label lbl1Letter 
      Caption         =   "One letter-based amino acid sequence"
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Tag             =   "6010"
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lbl3Letter 
      Caption         =   "Three letter-based amino acid sequence"
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Tag             =   "6020"
      Top             =   2160
      Width           =   1575
   End
End
Attribute VB_Name = "frmAminoAcidConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Purpose: Make sure keystrokes in an Amino-acid containing textbox are valid
Private Sub AminoAcidTextBoxKeyPressHandler(txtThisTextBox As TextBox, KeyAscii As Integer)
    ' Checks KeyAscii to see if it's valid
    ' If it isn't, it is set to 0
    
    If Not objMwtWin.IsModSymbol(Chr(KeyAscii)) Then
        Select Case KeyAscii
        Case 1          ' Ctrl+A -- select entire textbox
            txtThisTextBox.SelStart = 0
            txtThisTextBox.SelLength = Len(txtThisTextBox.Text)
            KeyAscii = 0
        Case 24, 3, 22  ' Cut, Copy, Paste is allowed
        Case 26
            ' Ctrl+Z = Undo
            KeyAscii = 0
            txtThisTextBox.Text = GetMostRecentTextBoxValue
        Case 8          ' Backspace key is allowed
        Case 32         ' Spaces are allowed
        Case 44, 46:    ' Decimal point (. or ,) is allowed (as a separator)
        Case 65 To 90, 97 To 122    ' Characters are allowed
        Case Else
            KeyAscii = 0
        End Select
    End If
    
End Sub

' Purpose: Convert a string of amino acid abbreviations to/from their 3 letter abbreviation
Private Sub ConvertSequence(bool1LetterTo3Letter As Boolean)
    ' If bool1LetterTo3Letter = True, then converting 1 letter codes to 3 letter codes
    ' If bool1LetterTo3Letter = False, then converting 3 letter codes to 1 letter codes
    
    Dim strWorkingSequence As String
    
    If bool1LetterTo3Letter Then
        ' Convert 1 letter symbols to 3 letter symbols
        strWorkingSequence = txt1LetterSequence.Text
        If Len(strWorkingSequence) = 0 Then Exit Sub
        
        objMwtWin.Peptide.SetSequence strWorkingSequence, ntgHydrogen, ctgHydroxyl, False
        
        txt3LetterSequence = objMwtWin.Peptide.GetSequence(True, False, cChkBox(chkSeparateWithDash), False, True)
    Else
        ' Convert 3 letter symbols to 1 letter symbols
        strWorkingSequence = txt3LetterSequence.Text
        If Len(strWorkingSequence) = 0 Then Exit Sub
        
        objMwtWin.Peptide.SetSequence strWorkingSequence
        
        txt1LetterSequence = objMwtWin.Peptide.GetSequence(False, cChkBox(chkSpaceEvery10), False, False, True)
    End If
End Sub

' Purpose: Copy 3 letter sequence to requested formula on frmMain
Private Sub CopySeqToFormula()
    Dim intTargetFormulaID As Integer, boolWindowWasHidden As Boolean
    Dim strSequence As String
    
    strSequence = txt3LetterSequence.Text
    If Len(strSequence) = 0 Then Exit Sub
    
    intTargetFormulaID = cboTargetFormula.ListIndex
    If intTargetFormulaID < 0 Then intTargetFormulaID = frmMain.GetTopFormulaIndex
    
    ' Need to make sure the form is shown so that adding a formula will work correctly
    ' Also necessary to assure the formula gets properly formatted
    boolWindowWasHidden = False
    If frmMain.Visible = False Then
        boolWindowWasHidden = True
        frmMain.WindowState = vbMinimized
        frmMain.Visible = True
    End If
    
    If intTargetFormulaID > frmMain.GetTopFormulaIndex And _
       frmMain.GetTopFormulaIndex < gMaxFormulaIndex Then
        frmMain.AddNewFormulaWrapper
    End If
    
    ' Use .SetSequence and .GetSequence to remove any modification symbols that might be present
    objMwtWin.Peptide.SetSequence strSequence, ntgHydrogen, ctgHydroxyl, True
    strSequence = objMwtWin.Peptide.GetSequence(True, False, cChkBox(chkSeparateWithDash), False, False)
    
    If Len(strSequence) = 0 Then
        MsgBox "Invalid sequence present.  Nothing to copy.", vbExclamation + vbOKOnly, "Error"
    Else
        frmMain.rtfFormula(intTargetFormulaID).Text = "H" & strSequence & "OH"
        frmMain.Calculate False, True, True, intTargetFormulaID, False, False, True, 1, True
    
        If boolWindowWasHidden Then
            frmMain.Visible = False
        End If
    End If
    
End Sub

' Purpose: Copy 3 letter sequence to the peptide sequence fragmentation modeller
Private Sub CopySeqToFragmentationModeller()
    frmFragmentationModelling.cboNotation.ListIndex = 1
    frmFragmentationModelling.PasteNewSequence txt3LetterSequence, True
    frmFragmentationModelling.Show
    
    If frmFragmentationModelling.WindowState <> vbMinimized Then
        frmFragmentationModelling.txtSequence.SetFocus
    End If

End Sub

' Purpose: Examines the formulas on frmMain, looking for the first blank one
Private Sub FindFirstAvailableFormula()
    ' Returns the highest formula (gMaxFormulaIndex) if all are in use
    
    Dim intIndex As Integer, intFirstAvailID As Integer
    
    With frmMain
        intFirstAvailID = -1
        For intIndex = 0 To frmMain.GetTopFormulaIndex
            If .rtfFormula(intIndex).Text = "" Then
                intFirstAvailID = intIndex
                Exit For
            End If
        Next intIndex
        If intFirstAvailID < 0 Then
            If frmMain.GetTopFormulaIndex < gMaxFormulaIndex Then
                intFirstAvailID = frmMain.GetTopFormulaIndex + 1
            Else
                intFirstAvailID = frmMain.GetTopFormulaIndex
            End If
        End If
    End With
    
    ' Now make sure this formula is highlighted in cboTargetFormula
    UpdateTargetFormulaCombo intFirstAvailID
    
End Sub

' Purpose: Resize this form and position and size controls
Private Sub ResizeForm(Optional boolEnlargeToMinimums As Boolean = False)
    Dim lngCmdConvertMidpoint As Long
    Dim boolSkipWidthAdjust As Boolean, boolSkipHeightAdjust As Boolean
    
    If Me.Width < 6300 Then
        If Not boolEnlargeToMinimums Then
            boolSkipWidthAdjust = True
        Else
            Me.Width = 6300
        End If
    End If
    
    If Me.Height < 4250 Then
        If Not boolEnlargeToMinimums Then
            boolSkipHeightAdjust = True
        Else
            Me.Height = 4250
        End If
    End If
    
    txt1LetterSequence.Top = 120
    If Not boolSkipHeightAdjust Then txt1LetterSequence.Height = Me.ScaleHeight * 0.2 - txt1LetterSequence.Top
    txt1LetterSequence.Left = 1920
    If Not boolSkipWidthAdjust Then txt1LetterSequence.Width = Me.ScaleWidth - txt1LetterSequence.Left - 280
    
    cmdConvertTo3Letter.Top = txt1LetterSequence.Top + txt1LetterSequence.Height + 120
    cmdConvertTo1Letter.Top = cmdConvertTo3Letter.Top
    
    lngCmdConvertMidpoint = txt1LetterSequence.Left + 0.5 * txt1LetterSequence.Width
    cmdConvertTo3Letter.Left = lngCmdConvertMidpoint - cmdConvertTo3Letter.Width - 240
    cmdConvertTo1Letter.Left = lngCmdConvertMidpoint + 120
    
    txt3LetterSequence.Top = cmdConvertTo1Letter.Top + cmdConvertTo1Letter.Height + 160
    If Not boolSkipHeightAdjust Then txt3LetterSequence.Height = Me.ScaleHeight - txt3LetterSequence.Top - 1000
    txt3LetterSequence.Left = txt1LetterSequence.Left
    txt3LetterSequence.Width = txt1LetterSequence.Width
    
    cmdCopy3Letter.Top = txt3LetterSequence.Top + txt3LetterSequence.Height + 120
    cboTargetFormula.Top = cmdCopy3Letter.Top + 180
    
    cmdClose.Top = cmdCopy3Letter.Top + 120
    cmdCopyToFragmentationModeller.Top = cmdCopy3Letter.Top
    
    lbl1Letter.Top = txt1LetterSequence.Top
    chkSpaceEvery10.Top = lbl1Letter.Top + 600
    
    lbl3Letter.Top = txt3LetterSequence.Top
    chkSeparateWithDash.Top = lbl3Letter.Top + 600

    If Not boolSkipWidthAdjust Then cmdClose.Left = Me.ScaleWidth - cmdClose.Width - 240
End Sub

' Purpose: Make sure cboTargetFormula has the same number of formulas as are displayed on frmMain
'            (plus 1 extra if frmMain.mTopFormulaIndex < gMaxFormulaIndex)
'          Also, set to the value of the first blank formula (or to frmMain.mTopFormulaIndex + 1 if not maxxed out)
Private Sub UpdateTargetFormulaCombo(Optional intFormulaIDToHighlight As Integer = -1)
    Dim intOldFormulaIDHighlighted As Integer, intIndex As Integer
    
    intOldFormulaIDHighlighted = cboTargetFormula.ListIndex
    With cboTargetFormula
        .Clear
        For intIndex = 0 To frmMain.GetTopFormulaIndex
            .AddItem CStr(intIndex + 1)
        Next intIndex
        If frmMain.GetTopFormulaIndex < gMaxFormulaIndex Then
            .AddItem CStr(frmMain.GetTopFormulaIndex + 2)
        End If
    End With
    
    If intFormulaIDToHighlight < 0 Then
        ' Simply make sure cboTargetFormula has the same number of formulas listed as frmMain
        cboTargetFormula.ListIndex = intOldFormulaIDHighlighted
    Else
        If intFormulaIDToHighlight < cboTargetFormula.ListCount Then
            cboTargetFormula.ListIndex = intFormulaIDToHighlight
        End If
    End If
End Sub

Private Sub chkSeparateWithDash_Click()
    ConvertSequence True
End Sub

Private Sub chkSpaceEvery10_Click()
    ConvertSequence False
End Sub

Private Sub cmdClose_Click()
    HideFormShowMain Me
End Sub

Private Sub cmdConvertTo1Letter_Click()
    ConvertSequence False
End Sub

Private Sub cmdConvertTo3Letter_Click()
    ConvertSequence True
End Sub

Private Sub cmdCopy3Letter_Click()
    CopySeqToFormula
End Sub

Private Sub cmdCopyToFragmentationModeller_Click()
    CopySeqToFragmentationModeller
End Sub

Private Sub Form_Activate()
    PossiblyHideMainWindow
    FindFirstAvailableFormula
End Sub

Private Sub Form_Load()
    SizeAndCenterWindow Me, cWindowUpperThird, 6000, 3500
    ResizeForm True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    QueryUnloadFormHandler Me, Cancel, UnloadMode
End Sub

Private Sub Form_Resize()
    ResizeForm False
End Sub

Private Sub txt1LetterSequence_GotFocus()
    HighlightOnFocus txt1LetterSequence
End Sub

Private Sub txt1LetterSequence_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        ConvertSequence True
    End If
End Sub

Private Sub txt1LetterSequence_KeyPress(KeyAscii As Integer)
    AminoAcidTextBoxKeyPressHandler txt1LetterSequence, KeyAscii
End Sub

Private Sub txt3LetterSequence_GotFocus()
    HighlightOnFocus txt3LetterSequence
End Sub

Private Sub txt3LetterSequence_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       ConvertSequence False
    End If
End Sub

Private Sub txt3LetterSequence_KeyPress(KeyAscii As Integer)
    AminoAcidTextBoxKeyPressHandler txt3LetterSequence, KeyAscii
End Sub
