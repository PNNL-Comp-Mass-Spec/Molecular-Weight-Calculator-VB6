VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MWCompoundClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

' Molecular Weight Calculator routines with ActiveX Class interfaces: MWCompoundClass
' Written by Matthew Monroe in Richland, WA

' The compound class can be used to represent a compound
' Use the Formula Property to enter the compound's formula
' Use ErrorDescription and CautionDescription to see if there are any problems with the formula
' Custom abbreviations can be defined using the SetAbbreviationInternal() function in ElementAndMassRoutines()
' Note that the standard amino acids and 16 other abbreviations are defined by default (see MemoryLoadAbbreviations())

' Use the Mass Property to get the mass of the compound

Private mStrFormula As String
Private mStrFormattedFormula As String

Private mValueForX As Double            ' The value to assign to x when present after a square bracket.
                                        ' For example, in C6H6[xBr] if x = 1, then the formula is treated like C6H6Br
                                        ' If x = 2, then the formula is treated like C6H6Br2

Private mCautionDescription As String
Private mErrorDescription As String
Private mErrorID As Long

Private mComputationStats As udtComputationStatsType

Public Function ConvertToEmpirical() As String
    ' Converts mStrFormula to its empirical formula and returns the result
    Dim strResult As String
    
    strResult = ConvertFormulaToEmpirical(mStrFormula)
    UpdateErrorAndCaution
    
    If mErrorDescription = "" Then
        mStrFormula = strResult
        mStrFormattedFormula = strResult
        ConvertToEmpirical = strResult
    Else
        ConvertToEmpirical = ErrorDescription
    End If
End Function

Public Function ElementPresent(intElementID As Integer) As Boolean
    ' Returns True if the element is present
    If intElementID >= 1 And intElementID <= ELEMENT_COUNT Then
        ElementPresent = mComputationStats.Elements(intElementID).Used
    Else
        ElementPresent = False
    End If
    
End Function

Public Function ExpandAbbreviations() As String
    ' Expands abbreviations in mStrFormula and returns the result

    Dim strResult As String
    
    strResult = ExpandAbbreviationsInFormula(mStrFormula)
    UpdateErrorAndCaution
    
    If mErrorDescription = "" Then
        mStrFormula = strResult
        mStrFormattedFormula = strResult
        ExpandAbbreviations = strResult
    Else
        ExpandAbbreviations = ErrorDescription
    End If

End Function

Public Function GetAtomCountForElement(ByVal intElementID As Integer) As Double
    ' Return the number of atoms of a given element that are present in the formula
    ' Note that the number of atoms is not necessarily an integer (e.g. C5.5)
    
    If intElementID >= 1 And intElementID <= ELEMENT_COUNT Then
        GetAtomCountForElement = mComputationStats.Elements(intElementID).Count
    Else
        GetAtomCountForElement = 0
    End If

End Function

Public Function GetPercentCompositionForElement(ByVal intElementID As Integer) As Double
    ' Returns the percent composition for element
    ' Returns -1 if an invalid ID
    
    If intElementID >= 1 And intElementID <= ELEMENT_COUNT Then
        GetPercentCompositionForElement = mComputationStats.PercentCompositions(intElementID).PercentComposition
    Else
        GetPercentCompositionForElement = -1
    End If

End Function

Public Function GetPercentCompositionForElementAsString(ByVal intElementID As Integer, Optional blnIncludeStandardDeviation As Boolean = True) As String
    ' Returns the percent composition and standard deviation for element
    ' Returns "" if an invalid ID
    Dim strElementSymbol As String
    Dim strPctComposition As String
    
    If intElementID >= 1 And intElementID <= ELEMENT_COUNT Then
        
        With mComputationStats.PercentCompositions(intElementID)
            
            strElementSymbol = GetElementSymbolInternal(intElementID) & ":"
            strPctComposition = ReturnFormattedMassAndStdDev(.PercentComposition, .StdDeviation, blnIncludeStandardDeviation, True)
            If .PercentComposition < 10 Then
                strPctComposition = " " & strPctComposition
            End If
            GetPercentCompositionForElementAsString = SpacePad(strElementSymbol, 4) & strPctComposition
        End With
    Else
        GetPercentCompositionForElementAsString = ""
    End If
    
End Function

Public Sub GetPercentCompositionForAllElements(ByRef strPctCompositionsOneBased() As String)
    ' Returns the percent composition for all elements in strPctCompositionsOneBased
    
    Dim intElementIndex As Integer, intMaxIndex As Integer
    
On Error GoTo GetPercentCompositionForAllElementsErrorHandler

    intMaxIndex = UBound(strPctCompositionsOneBased())
    
    If intMaxIndex < ELEMENT_COUNT Then
        ' Try to reserve more space in strPctCompositionsOneBased()
        On Error Resume Next
        ReDim strPctCompositionsOneBased(ELEMENT_COUNT)
        
        intMaxIndex = UBound(strPctCompositionsOneBased())
        
        On Error GoTo GetPercentCompositionForAllElementsErrorHandler
    End If
    
    If intMaxIndex >= ELEMENT_COUNT Then intMaxIndex = ELEMENT_COUNT
    
    ComputePercentComposition mComputationStats
   
    For intElementIndex = 1 To intMaxIndex
        With mComputationStats.PercentCompositions(intElementIndex)
            If .PercentComposition > 0 Then
                strPctCompositionsOneBased(intElementIndex) = ReturnFormattedMassAndStdDev(.PercentComposition, .StdDeviation)
            Else
                strPctCompositionsOneBased(intElementIndex) = ""
            End If
        End With
    Next intElementIndex
    
    Exit Sub
    
GetPercentCompositionForAllElementsErrorHandler:
    Debug.Print "Error occurred while copying percent composition values.  Probably an uninitialized array."
End Sub

Public Function GetUsedElementCount() As Integer
    ' Returns the number of unique elements present in mStrFormula
    
    Dim intTotalElements As Integer
    Dim intElementIndex As Integer

    ' Determine # of elements in formula
    intTotalElements = 0
    For intElementIndex = 1 To ELEMENT_COUNT
        ' Increment .TotalElements if element is present
        If mComputationStats.Elements(intElementIndex).Used Then
            intTotalElements = intTotalElements + 1
        End If
    Next intElementIndex
    
    GetUsedElementCount = intTotalElements
End Function

Public Function SetFormula(ByVal strNewFormula As String) As Long
    ' Provides an alternate method for setting the formula
    ' Returns ErrorID (0 if no error)
    
    Me.Formula = strNewFormula
    
    SetFormula = Me.ErrorID
End Function

Private Sub UpdateErrorAndCaution()
    mCautionDescription = GetCautionDescription()
    mErrorDescription = GetErrorDescription()
    mErrorID = GetErrorID()
End Sub

Private Sub UpdateMass()
    Dim dblMass As Double
    
    mStrFormattedFormula = mStrFormula
    
    ' mStrFormattedFormula is passed ByRef
    ' If gComputationOptions.CaseConversion = ccConvertCaseUp then mStrFormattedFormula is properly capitalized
    ' The mass of the compound is stored in mComputationStats.TotalMass
    dblMass = ParseFormulaPublic(mStrFormattedFormula, mComputationStats, False, mValueForX)
    
    ComputePercentComposition mComputationStats
    
    UpdateErrorAndCaution
End Sub

Public Function XIsPresentAfterBracket() As Boolean
    Dim intCharLoc As Integer
    
    If gComputationOptions.BracketsAsParentheses Then
        ' Treating brackets as parentheses, therefore an x after a bracket isn't allowed
        XIsPresentAfterBracket = False
    Else
        intCharLoc = InStr(LCase(mStrFormattedFormula), "[x")
        If intCharLoc > 0 Then
            If Mid(mStrFormattedFormula, intCharLoc + 1, 1) <> "e" Then
                XIsPresentAfterBracket = True
            Else
                XIsPresentAfterBracket = False
            End If
        Else
            XIsPresentAfterBracket = False
        End If
    End If
    
End Function

Public Property Get CautionDescription() As String
    CautionDescription = mCautionDescription
End Property

Public Property Let Charge(ByVal sngNewCharge As Single)
    mComputationStats.Charge = sngNewCharge
End Property

Public Property Get Charge() As Single
    Charge = mComputationStats.Charge
End Property

Public Property Get ErrorDescription() As String
    ErrorDescription = mErrorDescription
End Property

Public Property Get ErrorID() As Long
    ErrorID = mErrorID
End Property

Public Property Let Formula(ByVal strNewFormula As String)
    mStrFormula = strNewFormula
    
    ' Recompute the mass for this formula
    ' Updates Error and Caution statements if there is a problem
    UpdateMass
End Property

Public Property Get Formula() As String
    Formula = mStrFormula
End Property

Public Property Get FormulaCapitalized() As String
    FormulaCapitalized = mStrFormattedFormula
End Property

Public Property Get FormulaRTF() As String
    FormulaRTF = PlainTextToRtfInternal(Me.FormulaCapitalized, False)
End Property

Public Property Get Mass(Optional blnRecomputeMass As Boolean = True) As Double
    If blnRecomputeMass Then UpdateMass
    
    Mass = mComputationStats.TotalMass
End Property

Public Property Get MassAndStdDevString(Optional blnRecomputeMass As Boolean = True) As String
    If blnRecomputeMass Then UpdateMass
    
    With mComputationStats
        MassAndStdDevString = ReturnFormattedMassAndStdDev(.TotalMass, .StandardDeviation)
    End With
End Property

Public Property Get StandardDeviation() As Double
    StandardDeviation = mComputationStats.StandardDeviation
End Property

Public Property Let ValueForX(dblValueForX As Double)
    If dblValueForX >= 0 Then mValueForX = dblValueForX
End Property

Public Property Get ValueForX() As Double
    ValueForX = mValueForX
End Property

Private Sub Class_Initialize()
    mStrFormula = ""
    ValueForX = 1#
End Sub
