VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmAminoAcidModificationSymbols 
   Caption         =   "Amino Acid Modification Symbols Editor"
   ClientHeight    =   7095
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   10080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   10080
   StartUpPosition =   3  'Windows Default
   Tag             =   "15500"
   Begin VB.TextBox txtPhosphorylationSymbol 
      Height          =   300
      Left            =   2520
      TabIndex        =   11
      Text            =   "*"
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton cmdAddToList 
      Appearance      =   0  'Flat
      Caption         =   "&Add selected to list"
      Height          =   360
      Left            =   6600
      TabIndex        =   6
      Tag             =   "15810"
      Top             =   6600
      Width           =   2235
   End
   Begin VB.ListBox lstStandardModfications 
      Height          =   2985
      Left            =   5880
      TabIndex        =   5
      Top             =   3240
      Width           =   3975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   360
      Left            =   120
      TabIndex        =   7
      Tag             =   "4010"
      Top             =   6600
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   360
      Left            =   1320
      TabIndex        =   8
      Tag             =   "4020"
      Top             =   6600
      Width           =   1035
   End
   Begin VB.CommandButton cmdReset 
      Appearance      =   0  'Flat
      Caption         =   "Reset to &Defaults"
      Height          =   360
      Left            =   2520
      TabIndex        =   9
      Tag             =   "15800"
      ToolTipText     =   "Resets modification symbols to defaults"
      Top             =   6600
      Width           =   1935
   End
   Begin MSFlexGridLib.MSFlexGrid grdModSymbols 
      Height          =   3135
      Left            =   120
      TabIndex        =   3
      Tag             =   "15600"
      ToolTipText     =   "Click to change a modification symbol"
      Top             =   3285
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   5530
      _Version        =   393216
      Rows            =   17
      Cols            =   5
      FixedCols       =   0
      ScrollBars      =   2
   End
   Begin VB.Label lblPhosphorylationSymbolComment 
      Height          =   255
      Left            =   3600
      TabIndex        =   12
      Top             =   2190
      Width           =   2415
   End
   Begin VB.Label lblPhosphorylationSymbol 
      Caption         =   "Phosphorylation Symbol:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Tag             =   "15200"
      Top             =   2190
      Width           =   2295
   End
   Begin VB.Label lblModificationSymbolHeaderDirections 
      Caption         =   "Single click to edit or remove.  Click in blank row to add."
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Tag             =   "15710"
      Top             =   2880
      Width           =   5535
   End
   Begin VB.Label lblModificationSymbolHeader 
      Caption         =   "Modification Symbols Defined"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Tag             =   "15730"
      Top             =   2640
      Width           =   5535
   End
   Begin VB.Label Label1 
      Caption         =   "Standard Modifications"
      Height          =   255
      Left            =   5880
      TabIndex        =   4
      Tag             =   "15720"
      Top             =   2880
      Width           =   2655
   End
   Begin VB.Label lblExplanation 
      Caption         =   "Explanation"
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9735
   End
End
Attribute VB_Name = "frmAminoAcidModificationSymbols"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const FORM_WIDTH_MAX = 11000
Private Const FORM_HEIGHT_MAX = 7500

Private Const MODSYMBOLS_GRID_COL_COUNT = 4

Private Type udtModificationSymbolsType
    Symbol As String
    ModificationMass As Double
    IndicatesPhosphorylation As Boolean
    Comment As String
End Type

Private DefaultModSymbols() As udtModificationSymbolsType    ' 0-based array
Private DefaultModSymbolsCount As Long

Private SavedModSymbols() As udtModificationSymbolsType      ' 0-based array
Private SavedModSymbolsCount As Long

Private mValueChanged As Boolean
Private mDelayUpdate As Boolean

Private Sub AddDefaultModificationSymbol(strSymbol As String, dblMass As Double, blnIndicatesPhosphorylation As Boolean, strComment As String)
    
    With DefaultModSymbols(DefaultModSymbolsCount)
        .Symbol = strSymbol
        .ModificationMass = dblMass
        .IndicatesPhosphorylation = blnIndicatesPhosphorylation
        .Comment = strComment
    
        lstStandardModfications.AddItem .Symbol & ", " & .ModificationMass & ", " & .Comment
    End With
    DefaultModSymbolsCount = DefaultModSymbolsCount + 1
    
End Sub
Private Sub AddSelectedModificationToList()
    Dim lngIndex As Long, lngSelectedIndex As Long
    Dim lngErrorID As Long
    
    lngSelectedIndex = -1
    For lngIndex = 0 To lstStandardModfications.ListCount - 1
        If lstStandardModfications.Selected(lngIndex) Then
            lngSelectedIndex = lngIndex
            Exit For
        End If
    Next lngIndex
    
    If lngSelectedIndex >= 0 Then
        With DefaultModSymbols(lngSelectedIndex)
            ' The SetModificationSymbol() will update the symbol if it exists,
            '  or add a new one if it doesn't exist
            lngErrorID = objMwtWin.Peptide.SetModificationSymbol(.Symbol, .ModificationMass, .IndicatesPhosphorylation, .Comment)
            Debug.Assert lngErrorID = 0
        End With
    End If
    
    DisplayCurrentModificationSymbols
End Sub

Private Sub DisplayCurrentModificationSymbols()

    Dim lngIndex As Long, intCurrentRow As Integer, intCurrentCol As Integer
    Dim lngModSymbolCount As Long
    Dim strSymbol As String, strComment As String
    Dim dblMass As Double, blnPhosphorylation As Boolean
    Dim blnPhosphorylationSymbolFound As Boolean
    Dim lngError As Long
    
    If mDelayUpdate Then Exit Sub
    mDelayUpdate = True
    
    txtPhosphorylationSymbol = ""
    
    intCurrentRow = grdModSymbols.Row
    intCurrentCol = grdModSymbols.Col
    
    ' Grab the Modification Symbols from objMwtWin and fill the grid
    ' Place the first phosphorylation symbol found in txtPhosphorylationSymbol
    lngModSymbolCount = 0
    For lngIndex = 1 To objMwtWin.Peptide.GetModificationSymbolCount
         
        lngError = objMwtWin.Peptide.GetModificationSymbol(lngIndex, strSymbol, dblMass, blnPhosphorylation, strComment)
        Debug.Assert lngError = 0
        
        If blnPhosphorylation Then
            If Not blnPhosphorylationSymbolFound Then
                txtPhosphorylationSymbol = strSymbol
                blnPhosphorylationSymbolFound = True
            End If
        Else
            With grdModSymbols
                lngModSymbolCount = lngModSymbolCount + 1
                ' Enlarge grid if necessary
                If lngModSymbolCount >= .Rows - 1 Then
                    .Rows = lngModSymbolCount + 2
                End If
                
                .TextMatrix(lngModSymbolCount, 0) = lngIndex
                .TextMatrix(lngModSymbolCount, 1) = strSymbol
                .TextMatrix(lngModSymbolCount, 2) = Round(dblMass, 6)
                .TextMatrix(lngModSymbolCount, 3) = strComment
            End With
        End If
    Next lngIndex

    ' Erase remaining cells in grid
    For lngIndex = lngModSymbolCount + 1 To grdModSymbols.Rows - 1
        With grdModSymbols
            .TextMatrix(lngIndex, 0) = ""
            .TextMatrix(lngIndex, 1) = ""
            .TextMatrix(lngIndex, 2) = ""
            .TextMatrix(lngIndex, 3) = ""
        End With
    Next lngIndex
    
    ' Adjust row height of rows in grid
    For lngIndex = 1 To grdModSymbols.Rows - 1
        grdModSymbols.RowHeight(lngIndex) = TextHeight("123456789gT") + 60
    Next lngIndex
    
    ' Re-position cursor
    grdModSymbols.Row = intCurrentRow
    If intCurrentRow - 3 >= 1 Then
        grdModSymbols.TopRow = intCurrentRow - 3
    Else
        grdModSymbols.TopRow = 1
    End If
    grdModSymbols.Col = intCurrentCol

    mDelayUpdate = False
    
End Sub

Private Sub HandleGridClick()

    Dim intCurrentRow As Integer
    Dim strModSymbolID As String, lngModSymbolID As Long
    Dim strSymbol As String, strMass As String, strComment As String
    Dim strNewSymbol As String, strNewMass As String, strNewComment As String
    Dim lngErrorID As Long
    
    ' Determine the current row
    intCurrentRow = grdModSymbols.Row
    
    If intCurrentRow < 1 Then Exit Sub
    
    ' Determine which mod symbol the user clicked on
    strModSymbolID = grdModSymbols.TextMatrix(intCurrentRow, 0)
        
    If strModSymbolID = "" Then
        lngModSymbolID = 0
    Else
        lngModSymbolID = Val(strModSymbolID)
        With grdModSymbols
            strSymbol = .TextMatrix(intCurrentRow, 1)
            strMass = .TextMatrix(intCurrentRow, 2)
            strComment = .TextMatrix(intCurrentRow, 3)
        End With
    End If
    
    ' Display the dialog box and get user's response.
    With frmEditModSymbolDetails
        .lblHiddenButtonClickStatus = BUTTON_NOT_CLICKED_YET
        .txtSymbol = strSymbol
        .txtMass = strMass
        .txtComment = strComment
        
        .Show vbModal
        
        If .lblHiddenButtonClickStatus = BUTTON_NOT_CLICKED_YET Then .lblHiddenButtonClickStatus = BUTTON_CANCEL
    End With
            
    If Not frmEditModSymbolDetails.lblHiddenButtonClickStatus = BUTTON_CANCEL Then
        If frmEditModSymbolDetails.lblHiddenButtonClickStatus = BUTTON_RESET Then
            ' BUTTON_RESET indicates to remove the mod symbol
            If IsNumeric(strModSymbolID) Then
                lngModSymbolID = Val(strModSymbolID)
                
                lngErrorID = objMwtWin.Peptide.RemoveModificationByID(lngModSymbolID)
                Debug.Assert lngErrorID = 0
            End If
        Else
            With frmEditModSymbolDetails
                strNewSymbol = .txtSymbol
                strNewMass = .txtMass
                strNewComment = .txtComment
            End With
            
            If Len(strNewSymbol) > 0 Then
                If strNewSymbol <> strSymbol Then
                    ' User changed symbol; need to remove the old entry before adding the new one
                    objMwtWin.Peptide.RemoveModification (strSymbol)
                End If
                lngErrorID = objMwtWin.Peptide.SetModificationSymbol(strNewSymbol, CDblSafe(strNewMass), False, strNewComment)
                Debug.Assert lngErrorID = 0
            End If
        End If
       
        DisplayCurrentModificationSymbols
        mValueChanged = True
    End If

End Sub

Public Sub InitializeForm()
    Dim lngIndex As Long
    Dim lngErrorID As Long
    
    SavedModSymbolsCount = objMwtWin.Peptide.GetModificationSymbolCount
    
    ReDim SavedModSymbols(SavedModSymbolsCount)
    For lngIndex = 1 To SavedModSymbolsCount
        ' Note: SavedModSymbols() is 0-based, but the modifications in objMwtWin.Peptide are 1-based
        With SavedModSymbols(lngIndex - 1)
            lngErrorID = objMwtWin.Peptide.GetModificationSymbol(lngIndex, .Symbol, .ModificationMass, .IndicatesPhosphorylation, .Comment)
            Debug.Assert lngErrorID = 0
        End With
    Next lngIndex
    
    DisplayCurrentModificationSymbols
    
    mValueChanged = False
End Sub

Private Sub InitializeControls()
    Dim strMessage As String
    
    strMessage = LookupLanguageCaption(15550, "Modified residues in an amino acid sequence are notated by placing a modification symbol directly after the residue's 1 letter or 3 letter abbreviation.")
    strMessage = strMessage & "  " & LookupLanguageCaption(15555, "The phosphorylation symbol is defined separately because phosphorylated residues can lose a phosphate group during fragmentation.")
    strMessage = strMessage & "  " & LookupLanguageCaption(15560, "As an example, if the phosphorylation symbol is *, then in the sequence 'GlyLeuTyr*' the tyrosine residue is phosphorylated.")
    strMessage = strMessage & "  " & LookupLanguageCaption(15565, "Allowable symbols for user-defined modifications are") & " ! # $ % & ' * + ? ^ _ ` and ~"
    strMessage = strMessage & "  " & LookupLanguageCaption(15570, "Modification symbols can be more than 1 character long, though it is suggested that you keep them to just 1 character long to make it easier for the parsing routine to correctly parse residues with multiple modifications.")
    strMessage = strMessage & "  " & LookupLanguageCaption(15575, "If a residue has multiple modifications, then simply place the appropriate modification symbols after the residue, for example in 'FLE*#L' the E residue is modified by both the * and # modifications.")
    strMessage = strMessage & "  " & LookupLanguageCaption(15580, "Modification masses can be negative values, as well as positive values.")
    
    lblExplanation.Caption = strMessage

    With grdModSymbols
        .ColWidth(0) = 0
        .Cols = MODSYMBOLS_GRID_COL_COUNT
        .Rows = 2
        .Clear
        .TextMatrix(0, 0) = "ModSymbolID (Hidden)"
        .TextMatrix(0, 1) = LookupLanguageCaption(15610, "Symbol")
        .TextMatrix(0, 2) = LookupLanguageCaption(15620, "Mass")
        .TextMatrix(0, 3) = LookupLanguageCaption(15640, "Comment")
        
        .ColWidth(0) = 0
        .ColWidth(1) = 800
        .ColWidth(2) = 1000
        .ColWidth(3) = 3400
        
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignGeneral
        .ColAlignment(3) = flexAlignLeftCenter
        
        .Row = 1
        .Col = 1
    End With
    
    DisplayCurrentModificationSymbols

End Sub

Private Sub InitializeDefaultModSymbols()
    ' Initialize DefaultModSymbols()
    
    Dim lngIndex As Long
    Dim lngReturn As Long
    Dim strModSymbol As String, dblModMass As Double, blnIndicatesPhosphorylation As Boolean, strComment As String
    
    Dim objPeptideClass As New MWPeptideClass
    
    lstStandardModfications.Clear
    DefaultModSymbolsCount = 0
    ReDim DefaultModSymbols(objPeptideClass.GetModificationSymbolCount)
    
    For lngIndex = 1 To objPeptideClass.GetModificationSymbolCount
        lngReturn = objPeptideClass.GetModificationSymbol(lngIndex, strModSymbol, dblModMass, blnIndicatesPhosphorylation, strComment)

        If Not blnIndicatesPhosphorylation Then
            AddDefaultModificationSymbol strModSymbol, dblModMass, blnIndicatesPhosphorylation, strComment
        Else
            lblPhosphorylationSymbolComment = "[HPO3], " & Round(dblModMass, 3)
        End If
    Next lngIndex
    
    Set objPeptideClass = Nothing
    
End Sub

Private Sub RemoveAllPhosphorylationSymbols()
    Dim strSymbol As String
    Dim lngIndex As Long
    Dim blnIndicatesPhosphorylation As Boolean
    Dim lngErrorID As Long
    
    ' Remove all mod symbols with phosphorylation = True
    lngIndex = 1
    Do While lngIndex <= objMwtWin.Peptide.GetModificationSymbolCount
        lngErrorID = objMwtWin.Peptide.GetModificationSymbol(lngIndex, strSymbol, 0, blnIndicatesPhosphorylation, "")
        Debug.Assert lngErrorID = 0
        
        If blnIndicatesPhosphorylation = True Then
            lngErrorID = objMwtWin.Peptide.RemoveModificationByID(lngIndex)
        Else
            lngIndex = lngIndex + 1
        End If
    Loop

End Sub

Public Sub ResetModificationSymbolsToDefaults()
    objMwtWin.Peptide.SetDefaultModificationSymbols
    mValueChanged = True
    
    ' Update displayed mod list
    DisplayCurrentModificationSymbols
End Sub

Private Sub RestoreSavedModSymbols()
    Dim lngIndex As Long
    Dim lngErrorID As Long
    
    objMwtWin.Peptide.RemoveAllModificationSymbols
    
    For lngIndex = 1 To SavedModSymbolsCount
        ' Note: SavedModSymbols() is 0-based, but the modifications in objMwtWin.Peptide are 1-based
        With SavedModSymbols(lngIndex - 1)
            lngErrorID = objMwtWin.Peptide.SetModificationSymbol(.Symbol, .ModificationMass, .IndicatesPhosphorylation, .Comment)
            Debug.Assert lngErrorID = 0
        End With
    Next lngIndex
End Sub

Private Sub ValidateNewPhosphorylationSymbol()
    
    Const PHOSPHORYLATION_MASS As Double = 79.9663326
    
    Dim strNewSymbol As String
    Dim lngModSymbolID As Long
    Dim lngErrorID As Long
    Dim eResponse As VbMsgBoxResult
    
    If mDelayUpdate Then Exit Sub
    
    strNewSymbol = txtPhosphorylationSymbol
    
    If Len(strNewSymbol) = 0 Then
        RemoveAllPhosphorylationSymbols
        mValueChanged = True
    Else
        ' See if any of the current modification symbols match txtPhosphorylationSymbol
        ' If they do, warn user that they will be removed
        lngModSymbolID = objMwtWin.Peptide.GetModificationSymbolID(strNewSymbol)
        
        If lngModSymbolID > 0 Then
            eResponse = MsgBox(LookupLanguageCaption(15760, "Warning, the new phosphorylation symbol is already being used for another modification.  Do you really want to use this symbol for phosphorylation and consequently remove the other definition?"), vbQuestion + vbYesNoCancel + vbDefaultButton3, _
                               LookupLanguageCaption(15765, "Modification Symbol Conflict"))
        Else
            eResponse = vbYes
        End If
        
        If eResponse = vbYes Then
            RemoveAllPhosphorylationSymbols
            lngErrorID = objMwtWin.Peptide.SetModificationSymbol(strNewSymbol, PHOSPHORYLATION_MASS, True, "Phosphorylation")
            Debug.Assert lngErrorID = 0
            
            mValueChanged = True = True
        End If
    End If
    
    DisplayCurrentModificationSymbols
    
End Sub

Private Sub cmdAddToList_Click()
    AddSelectedModificationToList
End Sub

Private Sub cmdCancel_Click()
    Dim eResponse As VbMsgBoxResult

    If mValueChanged Then
        eResponse = YesNoBox(LookupLanguageCaption(15770, "Are you sure you want to lose all changes?"), _
                             LookupLanguageCaption(15775, "Closing Edit Modification Symbols Window"))
        If eResponse = vbYes Then
            RestoreSavedModSymbols
        Else
            Exit Sub
        End If
    End If
    
    mValueChanged = False
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    mValueChanged = False
    Me.Hide
End Sub

Private Sub cmdReset_Click()
    ResetModificationSymbolsToDefaults
End Sub

Private Sub Form_Load()
    SizeAndCenterWindow Me, cWindowUpperThird, FORM_WIDTH_MAX, FORM_HEIGHT_MAX
    
    InitializeDefaultModSymbols
    
    InitializeControls
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    QueryUnloadFormHandler Me, Cancel, UnloadMode
End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    With Me
        If .Height > FORM_HEIGHT_MAX Then .Height = FORM_HEIGHT_MAX
        If .Width > FORM_WIDTH_MAX Then .Width = FORM_WIDTH_MAX
    End With
    
End Sub

Private Sub grdModSymbols_Click()
    HandleGridClick
End Sub

Private Sub lstStandardModfications_DblClick()
    AddSelectedModificationToList
End Sub

Private Sub txtPhosphorylationSymbol_Change()
    Dim lngSelSave As Long
    
    With txtPhosphorylationSymbol
        lngSelSave = .SelStart
        ValidateNewPhosphorylationSymbol
        .SelStart = lngSelSave
    End With
End Sub

Private Sub txtPhosphorylationSymbol_KeyPress(KeyAscii As Integer)
    ModSymbolKeyPressHandler txtPhosphorylationSymbol, KeyAscii
End Sub
