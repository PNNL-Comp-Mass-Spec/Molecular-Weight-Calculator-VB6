VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmEditAbbrev 
   Caption         =   "Editing Abbreviations"
   ClientHeight    =   5730
   ClientLeft      =   2655
   ClientTop       =   1650
   ClientWidth     =   5565
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
   HelpContextID   =   1020
   Icon            =   "EDITABBR.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5730
   ScaleWidth      =   5565
   Tag             =   "9000"
   Begin MSFlexGridLib.MSFlexGrid grdAmino 
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Tag             =   "9140"
      ToolTipText     =   "Click to change an abbreviation"
      Top             =   120
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   3625
      _Version        =   393216
      Rows            =   17
      Cols            =   5
      FixedCols       =   0
      ScrollBars      =   2
   End
   Begin MSFlexGridLib.MSFlexGrid grdNormal 
      Height          =   2175
      Left            =   120
      TabIndex        =   1
      Tag             =   "9140"
      ToolTipText     =   "Click to change an abbreviation"
      Top             =   2400
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   3836
      _Version        =   393216
      Rows            =   17
      Cols            =   4
      FixedCols       =   0
      ScrollBars      =   2
   End
   Begin VB.CommandButton cmdReset 
      Appearance      =   0  'Flat
      Caption         =   "&Reset to Defaults"
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
      Left            =   3000
      TabIndex        =   2
      Tag             =   "9010"
      ToolTipText     =   "Resets abbreviations to the program defaults"
      Top             =   4680
      Width           =   1935
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
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
      Left            =   1800
      TabIndex        =   3
      Tag             =   "4020"
      Top             =   4680
      Width           =   1035
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
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
      Left            =   600
      TabIndex        =   4
      Tag             =   "4010"
      Top             =   4680
      Width           =   1035
   End
End
Attribute VB_Name = "frmEditAbbrev"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const AMINO_GRID_COL_COUNT = 7
Private Const ABBREV_GRID_COL_COUNT = 6

Private mValueChanged As Boolean

Private Sub HandleGridClick(ThisGrid As MSFlexGrid, blnAminoAcidGrid As Boolean)

    Dim strMessage As String
    
    Dim intCurrentRow As Integer
    Dim strAbbrevID As String, lngAbbrevID As Long
    Dim strSymbol As String, strFormula As String, strCharge As String, strOneLetterSymbol As String, strComment As String
    Dim strNewSymbol As String, strNewFormula As String, strNewCharge As String, strNewOneLetterSymbol As String, strNewComment As String
    Dim sngCharge As Single
    Dim lngErrorID As Long
    Dim lngMaxAbbrevCount As Long
    Dim dblTestMass As Double
    Dim dblTestMassOfFormula As Double
    
    ' Determine the current row
    intCurrentRow = ThisGrid.Row
    
    If intCurrentRow < 1 Then Exit Sub
    
    ' Determine which abbreviation the user clicked on
    strAbbrevID = ThisGrid.TextMatrix(intCurrentRow, 0)
        
    If strAbbrevID = "" Then
        lngAbbrevID = 0
        lngMaxAbbrevCount = objMwtWin.GetAbbreviationCountMax()
        If objMwtWin.GetAbbreviationCount >= lngMaxAbbrevCount Then
            MsgBox LookupLanguageCaption(9060, "Sorry, only " & Trim(lngMaxAbbrevCount) & " total abbreviations are allowed."), _
                   vbOKOnly & vbExclamation, LookupLanguageCaption(9050, "Maximum reached")
            Exit Sub
        End If
    Else
        lngAbbrevID = Val(strAbbrevID)
        With ThisGrid
            strSymbol = .TextMatrix(intCurrentRow, 1)
            strFormula = .TextMatrix(intCurrentRow, 2)
            strCharge = .TextMatrix(intCurrentRow, 3)
            If blnAminoAcidGrid Then
                strOneLetterSymbol = .TextMatrix(intCurrentRow, 4)
                strComment = .TextMatrix(intCurrentRow, 5)
            Else
                strComment = .TextMatrix(intCurrentRow, 4)
            End If
        End With
    End If
            
    ' The Message of the dialog box:
    strMessage = LookupLanguageCaption(9090, "The abbreviation or molecular formula will be changed to the value you type.  Select Remove to delete the abbreviation or Cancel to ignore any changes.")
    
    ' Display the dialog box and get user's response.
    With frmEditAbbrevDetails
        .lblHiddenButtonClickStatus = BUTTON_NOT_CLICKED_YET
        .lblInstructions.Caption = strMessage
        .txtSymbol = strSymbol
        .rtfFormula.Text = strFormula
        .txtCharge = strCharge
        .txtComment = strComment
        
        .txtOneLetterSymbol.Visible = blnAminoAcidGrid
        .lblOneLetterSymbol.Visible = blnAminoAcidGrid
        
        If blnAminoAcidGrid Then
            .txtOneLetterSymbol = strOneLetterSymbol
        Else
            .txtOneLetterSymbol = ""
        End If
                
        .Show vbModal
        
        If .lblHiddenButtonClickStatus = BUTTON_NOT_CLICKED_YET Then .lblHiddenButtonClickStatus = BUTTON_CANCEL
    End With
            
    If Not frmEditAbbrevDetails.lblHiddenButtonClickStatus = BUTTON_CANCEL Then
        If frmEditAbbrevDetails.lblHiddenButtonClickStatus = BUTTON_RESET Then
            ' BUTTON_RESET indicates to remove the abbreviation
            If IsNumeric(strAbbrevID) Then
                lngAbbrevID = Val(strAbbrevID)
                
                objMwtWin.RemoveAbbreviationByID lngAbbrevID
            End If
        Else
            With frmEditAbbrevDetails
                strNewSymbol = .txtSymbol
                strNewFormula = .rtfFormula.Text
                RemoveHeightAdjustChar strNewFormula
                strNewCharge = .txtCharge
                strNewOneLetterSymbol = .txtOneLetterSymbol
                strNewComment = .txtComment
            End With
            
            If IsNumeric(strNewCharge) Then
                sngCharge = Val(strNewCharge)
            Else
                sngCharge = 0
            End If
            
            lngErrorID = objMwtWin.SetAbbreviationByID(lngAbbrevID, strNewSymbol, strNewFormula, sngCharge, blnAminoAcidGrid, strNewOneLetterSymbol, strNewComment)
            
            If lngErrorID = 0 Then
                ' Check for circular reference
                dblTestMass = objMwtWin.ComputeMass(strNewSymbol)
                dblTestMassOfFormula = objMwtWin.ComputeMass(strNewFormula)
                If Round(dblTestMass, 3) <> Round(dblTestMassOfFormula, 3) Or dblTestMassOfFormula = 0 Then
                    ' Probably an invalid abbreviation
                    
                    MsgBox LookupLanguageCaption(9120, "Warning, test usage of the new/updated abbreviation did not give the same mass as the formula for the abbreviation.  This probably indicates a problem with the abbreviation's formula."), _
                           vbExclamation + vbOKOnly, LookupLanguageCaption(20350, "Error")
                End If
            End If
            
            If lngErrorID <> 0 Then MsgBox objMwtWin.GetMessageStatement(lngErrorID), vbInformation + vbOKOnly, LookupMessage(350)
        End If
       
        Me.MousePointer = vbHourglass
       
        ' Validate all of the abbreviations
        objMwtWin.ValidateAllAbbreviations

        Me.MousePointer = vbDefault

        UpdateGrid ThisGrid, blnAminoAcidGrid
        mValueChanged = True
    End If

End Sub

Private Sub InitializeGrids()
    With grdAmino
        .ColWidth(0) = 0
        .Cols = AMINO_GRID_COL_COUNT
        .TextMatrix(0, 0) = "AbbrevID (Hidden)"
        .TextMatrix(0, 1) = LookupLanguageCaption(9180, "Amino Acid Name")
        .TextMatrix(0, 2) = LookupLanguageCaption(9160, "Formula")
        .TextMatrix(0, 3) = LookupLanguageCaption(9150, "Charge")
        .TextMatrix(0, 4) = LookupLanguageCaption(9190, "1 Letter")
        .TextMatrix(0, 5) = LookupLanguageCaption(9195, "Comment")
        .TextMatrix(0, 6) = LookupLanguageCaption(20350, "Error")
        .Row = 1
        .Col = 1
    End With

    With grdNormal
        .ColWidth(0) = 0
        .Cols = ABBREV_GRID_COL_COUNT
        .TextMatrix(0, 0) = "AbbrevID (Hidden)"
        .TextMatrix(0, 1) = LookupLanguageCaption(9170, "Abbrev Name")
        .TextMatrix(0, 2) = LookupLanguageCaption(9160, "Formula")
        .TextMatrix(0, 3) = LookupLanguageCaption(9150, "Charge")
        .TextMatrix(0, 4) = LookupLanguageCaption(9195, "Comment")
        .TextMatrix(0, 5) = LookupLanguageCaption(20350, "Error")
        .Row = 1
        .Col = 1
    End With
    
    UpdateGrids

End Sub

Public Sub PositionFormControls()
    Dim tempHeight As Long
    
    With grdAmino
        .Left = 120
        .Top = 120
        .Width = frmEditAbbrev.Width - 300
        .Height = frmEditAbbrev.Height * 0.4
        .Cols = AMINO_GRID_COL_COUNT
        .ColWidth(0) = 0
        .ColWidth(1) = (.Width - 350) * 1 / 8
        .ColWidth(2) = (.Width - 350) * 1.8 / 8
        .ColWidth(3) = (.Width - 350) * 1 / 8
        .ColWidth(4) = (.Width - 350) * 1 / 8
        .ColWidth(5) = (.Width - 350) * 2.5 / 8
        .ColWidth(6) = (.Width - 350) * 0.7 / 8
        .ColAlignment(2) = vbLeftJustify
        .ColAlignment(3) = vbLeftJustify
        .ColAlignment(4) = vbLeftJustify
        .ColAlignment(5) = vbLeftJustify
        .ColAlignment(6) = vbLeftJustify
    End With
    
    With grdNormal
        .Left = grdAmino.Left
        .Top = grdAmino.Top + grdAmino.Height + 150
        .Width = grdAmino.Width
        tempHeight = frmEditAbbrev.Height - grdAmino.Height - 1400
        If tempHeight < 0 Then tempHeight = 0
        .Height = tempHeight
        .Cols = ABBREV_GRID_COL_COUNT
        .ColWidth(0) = 0
        .ColWidth(1) = (.Width - 350) * 1 / 7
        .ColWidth(2) = (.Width - 350) * 1.8 / 7
        .ColWidth(3) = (.Width - 350) * 1 / 7
        .ColWidth(4) = (.Width - 350) * 2.5 / 7
        .ColWidth(5) = (.Width - 350) * 0.7 / 7
        .ColAlignment(2) = vbLeftJustify
        .ColAlignment(3) = vbLeftJustify
        .ColAlignment(4) = vbLeftJustify
        .ColAlignment(5) = vbLeftJustify
    End With
    
    cmdOK.Top = grdNormal.Top + grdNormal.Height + 75
    cmdOK.Left = cmdOK.Left
    cmdCancel.Top = cmdOK.Top
    cmdReset.Top = cmdOK.Top
    
End Sub

Public Sub ResetValChangedToFalse()
    mValueChanged = False
End Sub

Private Sub UpdateGrid(ThisGrid As MSFlexGrid, blnAminoAcidGrid As Boolean)

    Dim lngIndex As Long, intCurrentRow As Integer, intCurrentCol As Integer
    Dim lngAbbrevCount As Long, lngThisAbbrevTypeCount As Integer
    Dim strSymbol As String, strFormula As String, strOneLetterSymbol As String
    Dim strComment As String, strErrorLabel As String
    Dim sngCharge As Single
    Dim blnIsAminoAcid As Boolean, blnInvalidSymbolOrFormula As Boolean
    Dim lngError As Long
    
    intCurrentRow = ThisGrid.Row
    intCurrentCol = ThisGrid.Col
    
    ' Grab the abbreviations from objMwtWin and fill the grid
    lngAbbrevCount = objMwtWin.GetAbbreviationCount
    
    lngThisAbbrevTypeCount = 0
    For lngIndex = 1 To lngAbbrevCount
        ' Grab data from objMwtWin
         
        lngError = objMwtWin.GetAbbreviation(lngIndex, strSymbol, strFormula, sngCharge, blnIsAminoAcid, strOneLetterSymbol, strComment, blnInvalidSymbolOrFormula)
        Debug.Assert lngError = 0
        
        If blnInvalidSymbolOrFormula Then
            strErrorLabel = LookupLanguageCaption(4060, "Yes")
        Else
            strErrorLabel = LookupLanguageCaption(4065, "No")
        End If
        
        ' The following assures that amino acids get shown in grdAmino and standard abbreviations get shown in grdNormal
        If blnAminoAcidGrid = blnIsAminoAcid Then
            lngThisAbbrevTypeCount = lngThisAbbrevTypeCount + 1
            
            With ThisGrid
                ' Enlarge grid if necessary
                If lngThisAbbrevTypeCount >= .Rows - 1 Then
                    .Rows = lngThisAbbrevTypeCount + 2
                End If
                
                .TextMatrix(lngThisAbbrevTypeCount, 0) = lngIndex
                .TextMatrix(lngThisAbbrevTypeCount, 1) = strSymbol
                .TextMatrix(lngThisAbbrevTypeCount, 2) = strFormula
                .TextMatrix(lngThisAbbrevTypeCount, 3) = Trim(Str(sngCharge))
                If .Cols >= AMINO_GRID_COL_COUNT And blnAminoAcidGrid Then
                    .TextMatrix(lngThisAbbrevTypeCount, 4) = strOneLetterSymbol
                    .TextMatrix(lngThisAbbrevTypeCount, 5) = strComment
                    .TextMatrix(lngThisAbbrevTypeCount, 6) = strErrorLabel
                Else
                    .TextMatrix(lngThisAbbrevTypeCount, 4) = strComment
                    .TextMatrix(lngThisAbbrevTypeCount, 5) = strErrorLabel
                End If
            End With
        End If
    Next lngIndex

    ' Erase remaining cells grid
    For lngIndex = lngThisAbbrevTypeCount + 1 To ThisGrid.Rows - 1
        With ThisGrid
            .TextMatrix(lngIndex, 0) = ""
            .TextMatrix(lngIndex, 1) = ""
            .TextMatrix(lngIndex, 2) = ""
            .TextMatrix(lngIndex, 3) = ""
                If .Cols >= AMINO_GRID_COL_COUNT And blnAminoAcidGrid Then
                    .TextMatrix(lngIndex, 4) = ""
                    .TextMatrix(lngIndex, 5) = ""
                    .TextMatrix(lngIndex, 6) = ""
                Else
                    .TextMatrix(lngIndex, 4) = ""
                    .TextMatrix(lngIndex, 5) = ""
                End If
        End With
    Next lngIndex
    
    ' Adjust row height of rows in grid
    For lngIndex = 1 To ThisGrid.Rows - 1
        ThisGrid.RowHeight(lngIndex) = TextHeight("123456789gT") + 60
    Next lngIndex
    
    ' Re-position cursor
    ThisGrid.Row = intCurrentRow
    If intCurrentRow - 3 >= 1 Then
        ThisGrid.TopRow = intCurrentRow - 3
    Else
        ThisGrid.TopRow = 1
    End If
    ThisGrid.Col = intCurrentCol

End Sub

Private Sub UpdateGrids()
    UpdateGrid grdAmino, True
    UpdateGrid grdNormal, False
End Sub

Private Sub cmdCancel_Click()
    Dim eResponse As VbMsgBoxResult

    If mValueChanged Then
        eResponse = YesNoBox(LookupLanguageCaption(9100, "Are you sure you want to lose all changes?"), _
                             LookupLanguageCaption(9105, "Closing Edit Abbreviations Box"))
        If eResponse = vbYes Then
            LoadAbbreviations False
        Else
            Exit Sub
        End If
    End If
    
    mValueChanged = False
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    
    Me.MousePointer = vbHourglass
    
    If mValueChanged Then
        SaveAbbreviations True, True
    End If
    
    mValueChanged = False
        
    Me.MousePointer = vbDefault
    Me.Hide
End Sub

Private Sub cmdReset_Click()
    Dim eResponse As VbMsgBoxResult

    eResponse = YesNoBox(LookupLanguageCaption(9110, "Are you sure you want to reset the abbreviations to the default abbreviations?"), _
                        LookupLanguageCaption(9115, "Reset to Defaults"))

    ' Evaluate the user's Response.
    If eResponse = vbYes Then
        mValueChanged = True
        LoadAbbreviations True
        
        UpdateGrids
    End If

End Sub

Private Sub Form_Activate()
    ' Put window in center of screen
    SizeAndCenterWindow Me, cWindowExactCenter, 8500, 6200
    UpdateGrids
End Sub

Private Sub Form_Load()
    
    PositionFormControls
    InitializeGrids
    
    mValueChanged = False
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
        cmdCancel_Click
    End If
End Sub

Private Sub Form_Resize()
    PositionFormControls
End Sub

Private Sub grdAmino_Click()
    HandleGridClick grdAmino, True
End Sub

Private Sub grdAmino_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then grdAmino_Click
End Sub

Private Sub grdNormal_Click()
    HandleGridClick grdNormal, False
End Sub

Private Sub grdNormal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then grdNormal_Click
End Sub
