Attribute VB_Name = "modSharedVBRoutines"
Option Explicit

'
' Collection of Visual Basic Functions
' Written by Matthew Monroe for use in applications
'
' First written in Chapel Hill, NC in roughly 2000
'
' Last Modified:    July 13, 2003
' Version:          1.59

Public KeyPressAbortProcess As Integer      ' Used with frmProgress
Public glbDecimalSeparator As String        ' Used to record whether the . or the , is the decimal point indicator (. in US while , in Europe)
Private mTextBoxValueSaved As String        ' Used to undo changes to a textbox

Private Const FIELD_DELIMETER = ","

' Constants for Centering Windows
Public Const cWindowExactCenter = 0
Public Const cWindowUpperThird = 1
Public Const cWindowLowerThird = 2
Public Const cWindowMiddleLeft = 3
Public Const cWindowMiddleRight = 4
Public Const cWindowTopCenter = 5
Public Const cWindowBottomCenter = 6
Public Const cWindowBottomRight = 7
Public Const cWindowBottomLeft = 8
Public Const cWindowTopRight = 9
Public Const cWindowTopLeft = 10

' The following function and constants are used to keep the application window
'   "on top" of other windows
'
Public Declare Function SetWindowPos Lib "user32" _
    (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
     ByVal x As Long, ByVal y As Long, ByVal cx As Long, _
     ByVal cy As Long, ByVal wFlags As Long) As Long

' Set some constant values (from WIN32API.TXT).
Private Const conHwndTopmost = -1
Private Const conHwndNoTopmost = -2
Private Const conSwpNoActivate = &H10
Private Const conSwpShowWindow = &H40

'' Used for Internet Access
''' Used in the GetUrlSource() Function
'''   from http://www.planet-source-code.com/vb/scripts/ShowCode.asp?lngWId=1&txtCodeId=24465
''Public Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
''Public Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" (ByVal hInternetSession As Long, ByVal sURL As String, ByVal sHeaders As String, ByVal lHeadersLength As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
''Public Declare Function InternetReadFile Lib "wininet.dll" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
''Public Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer
''
''Public Const IF_FROM_CACHE = &H1000000
''Public Const IF_MAKE_PERSISTENT = &H2000000
''Public Const IF_NO_CACHE_WRITE = &H4000000
''
''
'''used for shelling out to the default web browser
''Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
''Public Const conSwNormal = 1
''
''Private Const BUFFER_LEN = 2048


' Return the handle of a window given its name or class.  Pass only one of the parameters,
'  using vbNullString for the other.
'
' For example: hwnd = FindWindow(vbNullString, "My Window Caption")
'
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetActiveWindow Lib "user32" () As Long
Public Declare Function GetTopWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long

' GetWindow() Constants
Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDLAST = 1
Public Const GW_HWNDNEXT = 2
Public Const GW_HWNDPREV = 3
Public Const GW_OWNER = 4
Public Const GW_CHILD = 5
Public Const GW_MAX = 5
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long

'
' More Window Functions
'
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function GetNextWindow Lib "user32" Alias "GetWindow" (ByVal hwnd As Long, ByVal wFlag As Long) As Long

' Functions for finding the size of the desktop

Private Type Rect
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Type POINTAPI
        x As Long
        y As Long
End Type

Private Type WINDOWPLACEMENT
        Length As Long
        flags As Long
        showCmd As Long
        ptMinPosition As POINTAPI
        ptMaxPosition As POINTAPI
        rcNormalPosition As Rect
End Type

Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetWindowPlacement Lib "user32" (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long

' Other functions
'
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' Api Functions
Public Declare Function GetTickCount Lib "kernel32" () As Long

' Functions for selecting a directory (Function BrowseForFolder)

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal hMem As Long)
Private Declare Function GetOpenFileName Lib "comdlg32" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

Private Type BROWSEINFO
    hwndOwner      As Long
    pidlRoot       As Long
    pszDisplayName As Long
    lpszTitle      As Long
    ulFlags        As Long
    lpfnCallback   As Long
    lParam         As Long
    iImage         As Long
End Type

'used with GetOpenFileName function
Private Type OPENFILENAME
        lStructSize As Long
        hwndOwner As Long
        hInstance As Long
        lpstrFilter As String
        lpstrCustomFilter As String
        nMaxCustFilter As Long
        nFilterIndex As Long
        lpstrFile As String
        nMaxFile As Long
        lpstrFileTitle As String
        nMaxFileTitle As Long
        lpstrInitialDir As String
        lpstrTitle As String
        flags As Long
        nFileOffset As Integer
        nFileExtension As Integer
        lpstrDefExt As String
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
End Type

' Possibly, this should be MAX_PATH = 257
Private Const MAX_PATH = 512

Private Const WM_USER = &H400
Private Const BFFM_INITIALIZED = 1
Private Const BFFM_SELCHANGED = 2
Private Const BFFM_SETSTATUSTEXT = (WM_USER + 100)
Private Const BFFM_SETSELECTION = (WM_USER + 102)

Private Const BIF_RETURNONLYFSDIRS = &H1
Private Const BIF_DONTGOBELOWDOMAIN = &H2
Private Const BIF_STATUSTEXT = &H4
Private Const BIF_RETURNFSANCESTORS = &H8
Private Const BIF_EDITBOX = &H10
Private Const BIF_NEWDIALOGSTYLE = &H20
Private Const BIF_USENEWUI = &H40
Private Const BIF_INCLUDECOMPUTERS = &H1000
Private Const BIF_INCLUDEPRINTERS = &H2000
Private Const BIF_INCLUDEFILES = &H4000

'OPENFILENAME structure flags constants (used with Open & Save dialogs)
Private Const OFN_ALLOWMULTISELECT = &H200
Private Const OFN_CREATEPROMPT = &H2000
Private Const OFN_ENABLEHOOK = &H20
Private Const OFN_ENABLETEMPLATE = &H40
Private Const OFN_ENABLETEMPLATEHANDLE = &H80
Private Const OFN_EXPLORER = &H80000
Private Const OFN_EXTENSIONDIFFERENT = &H400
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_LONGNAMES = &H200000
Private Const OFN_NOCHANGEDIR = &H8
Private Const OFN_NODEREFERENCELINKS = &H100000
Private Const OFN_NOLONGNAMES = &H40000
Private Const OFN_NONETWORKBUTTON = &H20000
Private Const OFN_NOREADONLYRETURN = &H8000
Private Const OFN_NOTESTFILECREATE = &H10000
Private Const OFN_NOVALIDATE = &H100
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_PATHMUSTEXIST = &H800
Private Const OFN_READONLY = &H1
Private Const OFN_SHAREAWARE = &H4000
Private Const OFN_SHAREFALLTHROUGH = 2
Private Const OFN_SHARENOWARN = 1
Private Const OFN_SHAREWARN = 0
Private Const OFN_SHOWHELP = &H10
Private Const OFS_MAXPATHNAME = 512

Private mCurrentDirectory As String   ' The last directory selected using BrowseForFolder()
'

Public Sub AppendToString(ByRef strThisString As String, strAppendText As String, Optional blnAddCarriageReturn As Boolean = False)
    strThisString = strThisString & strAppendText
    
    If blnAddCarriageReturn Then
        strThisString = strThisString & vbCrLf
    End If
End Sub

Public Function AssureNonZero(lngNumber As Long) As Long
    ' Returns a non-zero number, either -1 if lngNumber = 0 or lngNumber if it's nonzero
    If lngNumber = 0 Then
        AssureNonZero = -1
    Else
        AssureNonZero = lngNumber
    End If
End Function

Public Function BrowseForFileOrFolder(ByVal lngOwnderhwnd As Long, Optional ByRef strStartPath As String, Optional ByVal strTitle As String = "Select File", Optional blnReturnFoldersOnly As Boolean = False, Optional strFilterDescription As String = "All") As String
    ' Returns the path to the selected file or folder
    ' Returns "" if cancelled
    
    '=====================================================================================
    ' Browse for a Folder using SHBrowseForFolder API function with a callback
    ' function BrowseCallbackProc.
    '
    ' This Extends the functionality that was given in the
    ' MSDN Knowledge Base article Q179497 "HOWTO: Select a Directory
    ' Without the Common Dialog Control".
    '
    ' After reading the MSDN knowledge base article Q179378 "HOWTO: Browse for
    ' Folders from the Current Directory", I was able to figure out how to add
    ' a callback function that sets the starting directory and displays the
    ' currently selected path in the "Browse For Folder" dialog.
    '
    ' I used VB 6.0 (SP3) to compile this code.  Should work in VB 5.0.
    ' However, because it uses the AddressOf operator this code will not
    ' work with versions below 5.0.
    '
    ' This code works in Window 95a so I assume it will work with later versions.
    '
    ' Stephen Fonnesbeck
    ' steev@xmission.com
    ' http://www.xmission.com/~steev
    ' Feb 20, 2000
    '
    '=====================================================================================
    ' Usage:
    '
    '    Dim folder As String
    '    folder = BrowseForFileOrFolder(Me, "C:\startdir\anywhere", True, "Select a Directory")
    '    If Len(folder) = 0 Then Exit Sub  'User Selected Cancel
    '
    '=====================================================================================
    '
    '
    ' Code extended by Matthew Monroe to also allow selection of files, using an example
    ' from http://www.thescarms.com/VBasic/DirectoryBrowser.asp
    '
    
    Dim lpIDList As Long
    Dim sBuffer As String
    Dim tBrowseInfo As BROWSEINFO
    
    If Len(strStartPath) > 0 Then
        mCurrentDirectory = strStartPath & vbNullChar
    End If
    
    If Len(mCurrentDirectory) = 0 Then mCurrentDirectory = vbNullChar
    
    With tBrowseInfo
        .hwndOwner = lngOwnderhwnd
        .lpszTitle = lstrcat(strTitle, "")
        .ulFlags = BIF_RETURNONLYFSDIRS Or BIF_DONTGOBELOWDOMAIN Or BIF_STATUSTEXT Or BIF_EDITBOX Or BIF_NEWDIALOGSTYLE Or BIF_USENEWUI
        If Not blnReturnFoldersOnly Then
            .ulFlags = .ulFlags Or BIF_INCLUDEFILES
        End If
        .lpfnCallback = GetAddressofFunction(AddressOf BrowseCallbackProc)  'get address of function.
    End With
    
    lpIDList = SHBrowseForFolder(tBrowseInfo)
    If (lpIDList) Then
        sBuffer = Space(MAX_PATH)
        SHGetPathFromIDList lpIDList, sBuffer
        sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
      
        Call CoTaskMemFree(lpIDList)
      
        BrowseForFileOrFolder = sBuffer
        strStartPath = sBuffer
    Else
        BrowseForFileOrFolder = ""
    End If
 
End Function

' Used with BrowseForFolder
Private Function BrowseCallbackProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal lp As Long, ByVal pData As Long) As Long
 
    Dim lpIDList As Long
    Dim Ret As Long
    Dim sBuffer As String
    
    On Error Resume Next  'Sugested by MS to prevent an error from propagating back into the calling process.
    
    Select Case uMsg
    Case BFFM_INITIALIZED
        Call SendMessage(hwnd, BFFM_SETSELECTION, 1, mCurrentDirectory)
    
    Case BFFM_SELCHANGED
        sBuffer = Space(MAX_PATH)
    
        Ret = SHGetPathFromIDList(lp, sBuffer)
        If Ret = 1 Then
            Call SendMessage(hwnd, BFFM_SETSTATUSTEXT, 0, sBuffer)
        End If

    End Select
    
    BrowseCallbackProc = 0
 
End Function

' This function allows you to assign a function pointer to a vaiable.
' Used with BrowseForFolder
Private Function GetAddressofFunction(add As Long) As Long
    GetAddressofFunction = add
End Function

Public Function cChkBox(value As Integer) As Boolean
    ' Converts a checkbox value to true or false
    If value = vbChecked Then
        cChkBox = True
    Else
        cChkBox = False
    End If
End Function

Public Function CLngRoundUp(dblValue As Double) As Long
    Dim lngNewValue As Long
    lngNewValue = CLng(dblValue)
    
    If lngNewValue <> dblValue Then
        lngNewValue = lngNewValue + 1
    End If
    
    CLngRoundUp = lngNewValue
End Function

Public Function CLngSafe(strValue As String) As Long
    On Error Resume Next
    If IsNumeric(strValue) Then
        CLngSafe = CLng(strValue)
    End If
End Function

Public Function CBoolSafe(strValue As String) As Boolean
    On Error GoTo ReturnFalse
    CBoolSafe = CBool(Trim(strValue))
    Exit Function
ReturnFalse:
    CBoolSafe = False
End Function

Public Function CDblSafe(strWork As String) As Double
    On Error Resume Next
    If IsNumeric(strWork) Then
        CDblSafe = CDbl(strWork)
    End If
End Function

Public Function CIntSafeDbl(dblWork As Double) As Integer
    If dblWork <= 32767 And dblWork >= -32767 Then
        CIntSafeDbl = CInt(dblWork)
    Else
        If dblWork < 0 Then
            CIntSafeDbl = -32767
        Else
            CIntSafeDbl = 32767
        End If
    End If
End Function

Public Function CIntSafe(strWork As String) As Integer
    If IsNumeric(strWork) Then
        CIntSafe = CIntSafeDbl(CDbl(strWork))
    ElseIf LCase(strWork) = "true" Then
        CIntSafe = -1
    Else
        CIntSafe = 0
    End If
End Function

Public Function Combinatorial(a As Integer, B As Integer) As Single
    If a > 170 Or B > 170 Then
        Debug.Print "Cannot compute factorial of a number over 170.  Thus, cannot compute the combination."
        Combinatorial = -1
    ElseIf a < B Then
        Debug.Print "First number should be greater than or equal to the second number"
        Combinatorial = -1
    Else
        Combinatorial = Factorial(a) / (Factorial(B) * Factorial(a - B))
    End If
End Function

Public Function CompactPathString(ByVal strPathToCompact As String, Optional ByVal lngMaxLength As Long = 40) As String
    ' Recursive function to shorten strPathToCompact to a maximum length of lngMaxLength
    
    ' The following is example output
    ' Note that when drive letters or subdirectories are present, the a minimum length is imposed
    ' For "C:\My Documents\Readme.txt"
    '   Minimum string returned=  C:\M..\Rea..
    '   Length for 20 characters= C:\My D..\Readme.txt
    '   Length for 25 characters= C:\My Docume..\Readme.txt
    '
    ' For "C:\My Documents\Word\Business\Finances.doc"
    '   Minimum string returned=  C:\...\B..\Fin..
    '   Length for 20 characters= C:\...\B..\Finance..
    '   Length for 25 characters= C:\...\Bus..\Finances.doc
    '   Length for 32 characters= C:\...\W..\Business\Finances.doc
    '   Length for 40 characters= C:\My Docum..\Word\Business\Finances.doc
    
    Dim strPath(4) As String        ' 0-based array
    Dim intPartCount As Integer
    
    Dim strLeadingChars As String
    Dim strShortenedPath As String
    
    Dim lngCharLoc As Long
    Dim intLoopCount As Integer, intFileNameIndex As Integer
    Dim lngShortLength As Long, lngOverLength As Long
    Dim lngLeadingCharsLength As Long
    Dim intMultiPathCorrection As Integer
    
    If lngMaxLength < 3 Then lngMaxLength = 3
    
    ' Determine the name of the first directory following the leading ?:\ or \\
    
    strPathToCompact = Trim(strPathToCompact)
    
    If Len(strPathToCompact) <= lngMaxLength Then
        CompactPathString = strPathToCompact
        Exit Function
    End If
    
    intPartCount = 1
    If Left(strPathToCompact, 2) = "\\" Then
        strLeadingChars = "\\"
        lngCharLoc = InStr(3, strPathToCompact, "\")
        If lngCharLoc > 0 Then
            strLeadingChars = "\\" & Mid(strPathToCompact, 3, lngCharLoc - 2)
            strPath(0) = Mid(strPathToCompact, lngCharLoc + 1)
        Else
            strPath(0) = Mid(strPathToCompact, 3)
        End If
    ElseIf Left(strPathToCompact, 2) = ".\" Then
        strLeadingChars = Left(strPathToCompact, 2)
        strPath(0) = Mid(strPathToCompact, 3)
    ElseIf Left(strPathToCompact, 3) = "..\" Or Mid(strPathToCompact, 2, 2) = ":\" Then
        strLeadingChars = Left(strPathToCompact, 3)
        strPath(0) = Mid(strPathToCompact, 4)
    Else
        strPath(0) = strPathToCompact
    End If
    
    ' Examine strPath(0) to see if there are 1, 2, or more subdirectories
    intLoopCount = 0
    Do
        lngCharLoc = InStr(strPath(intPartCount - 1), "\")
        If lngCharLoc > 0 Then
            strPath(intPartCount) = Mid(strPath(intPartCount - 1), lngCharLoc + 1)
            strPath(intPartCount - 1) = Left(strPath(intPartCount - 1), lngCharLoc)
            intPartCount = intPartCount + 1
        End If
        intLoopCount = intLoopCount + 1
    Loop While intLoopCount < 3
    
    If intPartCount = 1 Then
        ' No \ found, we're forced to shorten the filename (though if a UNC, then can shorten part of the UNC)
        
        If Left(strLeadingChars, 2) = "\\" Then
            lngLeadingCharsLength = Len(strLeadingChars)
            If lngLeadingCharsLength > 5 Then
                ' Can shorten the server name as needed
                lngShortLength = lngMaxLength - Len(strPath(0)) - 3
                If lngShortLength < lngLeadingCharsLength Then
                    If lngShortLength < 3 Then lngShortLength = 3
                    strLeadingChars = Left(strLeadingChars, lngShortLength) & "..\"
                End If
                
            End If
        End If
        
        lngShortLength = lngMaxLength - Len(strLeadingChars) - 2
        If lngShortLength < 3 Then lngShortLength = 3
        If lngShortLength < Len(strPath(0)) - 2 Then
            strShortenedPath = strLeadingChars & Left(strPath(0), lngShortLength) & ".."
        Else
            strShortenedPath = strLeadingChars & strPath(0)
        End If
    Else
        ' Found one (or more) subdirectories
        
        ' First check if strPath(1) = "...\"
        If strPath(0) = "...\" Then
            intMultiPathCorrection = 4
            strPath(0) = strPath(1)
            strPath(1) = strPath(2)
            strPath(2) = strPath(3)
            strPath(3) = ""
            intPartCount = 3
        Else
            intMultiPathCorrection = 0
        End If
        
        ' Shorten the first to as little as possible
        ' If not short enough, replace the first with ... and call this function again
        lngShortLength = lngMaxLength - Len(strLeadingChars) - Len(strPath(3)) - Len(strPath(2)) - Len(strPath(1)) - 3 - intMultiPathCorrection
        If lngShortLength < 1 And Len(strPath(2)) > 0 Then
            ' Not short enough, but other subdirectories are present
            ' Thus, can call this function recursively
            strShortenedPath = strLeadingChars & "...\" & strPath(1) & strPath(2) & strPath(3)
            strShortenedPath = CompactPathString(strShortenedPath, lngMaxLength)
        Else
            If Left(strLeadingChars, 2) = "\\" Then
                lngLeadingCharsLength = Len(strLeadingChars)
                If lngLeadingCharsLength > 5 Then
                    ' Can shorten the server name as needed
                    lngShortLength = lngMaxLength - Len(strPath(3)) - Len(strPath(2)) - Len(strPath(1)) - 7 - intMultiPathCorrection
                    If lngShortLength < lngLeadingCharsLength - 3 Then
                        If lngShortLength < 3 Then lngShortLength = 3
                        strLeadingChars = Left(strLeadingChars, lngShortLength) & "..\"
                    End If
                    
                    ' Recompute lngShortLength
                    lngShortLength = lngMaxLength - Len(strLeadingChars) - Len(strPath(3)) - Len(strPath(2)) - Len(strPath(1)) - 3 - intMultiPathCorrection
                End If
            End If
            
            If intMultiPathCorrection > 0 Then
                strLeadingChars = strLeadingChars & "...\"
            End If
            
            If lngShortLength < 1 Then lngShortLength = 1
            strPath(0) = Left(strPath(0), lngShortLength) & "..\"
            strShortenedPath = strLeadingChars & strPath(0) & strPath(1) & strPath(2) & strPath(3)
        
            ' See if still to long
            ' If it is, then will need to shorten the filename too
            lngOverLength = Len(strShortenedPath) - lngMaxLength
            If lngOverLength > 0 Then
                ' Need to shorten filename too
                ' Determine which index the filename is in
                For intFileNameIndex = intPartCount - 1 To 0 Step -1
                    If Len(strPath(intFileNameIndex)) > 0 Then Exit For
                Next intFileNameIndex
                
                lngShortLength = Len(strPath(intFileNameIndex)) - lngOverLength - 2
                If lngShortLength < 4 Then
                    strPath(intFileNameIndex) = Left(strPath(intFileNameIndex), 3) & ".."
                Else
                    strPath(intFileNameIndex) = Left(strPath(intFileNameIndex), lngShortLength) & ".."
                End If
                
                strShortenedPath = strLeadingChars & strPath(0) & strPath(1) & strPath(2) & strPath(3)
            End If
        
        End If
    End If
    
    CompactPathString = strShortenedPath
End Function

Public Function TestCompactFileName(strTestString As String)
    Dim lngIndex As Long
    
    For lngIndex = 15 To Len(strTestString) + 2
        Debug.Print Format(lngIndex, "000") & ": " & CompactPathString(strTestString, lngIndex)
    Next lngIndex
End Function

Private Function ComputeAverage(ByRef sngArrayZeroBased() As Single, lngArrayCount As Long, Optional ByRef sngMedian As Single, Optional ByRef sngMinimum As Single, Optional ByRef sngMaximum As Single) As Single
    ' Computes the average and returns it
    ' Computes other stats and returns them

    Dim lngIndex As Long
    Dim sngSumForAveraging As Single, lclAverage As Single, lclMedian As Single
    Dim lclMinimum As Single, lclMaximum As Single
    
    If lngArrayCount = 0 Then
        lclAverage = 0
        lclMedian = 0
    Else
        lclMinimum = sngArrayZeroBased(0)
        lclMaximum = lclMinimum
        
        sngSumForAveraging = 0
        For lngIndex = 0 To lngArrayCount - 1
            sngSumForAveraging = sngSumForAveraging + sngArrayZeroBased(lngIndex)
            
            If sngArrayZeroBased(lngIndex) < lclMinimum Then
                lclMinimum = sngArrayZeroBased(lngIndex)
            End If
            
            If sngArrayZeroBased(lngIndex) > lclMaximum Then
                lclMaximum = sngArrayZeroBased(lngIndex)
            End If
            
        Next lngIndex
        
        lclAverage = sngSumForAveraging / lngArrayCount
        
        lclMedian = ComputeMedian(sngArrayZeroBased(), lngArrayCount)
    End If
    
    
    sngMinimum = lclMinimum
    sngMaximum = lclMaximum

    sngMedian = lclMedian
    ComputeAverage = lclAverage
    
End Function

Public Function ComputeMedian(ByRef sngArrayZeroBased() As Single, lngArrayCount As Long, Optional blnRigorousMedianForEvenNumberedDatasets As Boolean = True) As Single
    ' Finds the median value in sngIniputArrayZeroBased() using the Select() function from Numerical Recipes in C
    ' Returns the median
    
    ' If lngArrayCount is Odd, then calls Select, (N+1)/2 for lngElementToSelect
    ' If lngArrayCount is Even, and blnRigorousMedianForEvenNumberedDatasets is True, or lngArrayCount < 100,
    '  then calls the select function twice, grabbing element N/2 and N/2 + 1
    ' Otherwise, if lngArrayCount is Even, but blnRigorousMedianForEvenNumberedDatasets (and lngArrayCount >= 100)
    '  then simply calls the function one, grabbing element N/2 and using this as an approximate median value
    
    Dim lngIndex As Long, lngElementToGrab As Long
    Dim blnEvenCount As Boolean
    Dim sngInputArrayOneBased() As Single
    Dim sngMedianValue As Single, sngMedianValueNextLarger As Single
    
    If lngArrayCount <= 0 Then Exit Function
    
    If lngArrayCount = 1 Then
        ComputeMedian = sngArrayZeroBased(0)
        Exit Function
    End If
    
    ' Since the Select function requires the array be one-based, I make a copy of the data in the Zero Based array
    ' This also prevents sngArrayZeroBased() from being disturbed
    
    ReDim sngInputArrayOneBased(lngArrayCount)
    
    For lngIndex = 0 To lngArrayCount - 1
        sngInputArrayOneBased(lngIndex + 1) = sngArrayZeroBased(lngIndex)
    Next lngIndex
    
    If lngArrayCount / 2# = CLng(lngArrayCount / 2#) Then
        ' lngArrayCount is even
        blnEvenCount = True
    Else
        ' lngArrayCount is odd
        blnEvenCount = False
    End If
    
    If blnEvenCount Then
        lngElementToGrab = lngArrayCount / 2
        sngMedianValue = SelectValue(sngInputArrayOneBased(), lngArrayCount, lngElementToGrab)
        
        If blnRigorousMedianForEvenNumberedDatasets Or lngArrayCount < 100 Then
            sngMedianValueNextLarger = SelectValue(sngInputArrayOneBased(), lngArrayCount, lngElementToGrab + 1)
            sngMedianValue = (sngMedianValue + sngMedianValueNextLarger) / 2
        End If
    Else
        lngElementToGrab = (lngArrayCount + 1) / 2
        sngMedianValue = SelectValue(sngInputArrayOneBased(), lngArrayCount, lngElementToGrab)
    End If
    
    ComputeMedian = sngMedianValue
    
End Function

Public Function ConstructFormatString(ByVal dblThisValue As Double, Optional ByRef intDigitsInFormattedValue As Integer) As String
    ' Examines dblThisValue and constructs a format string based on its magnitude
    ' For example, dblThisValue = 1234 will return "0"
    '              dblThisValue = 2.4323 will return "0.0000"
    '
    ' In addition, returns the length of the string representation of dblThisValue
    '  using the determined format string
    Dim lngExponentValue As Long, intDigitsInLabel As Integer
    Dim strWork As String, strFormatString As String
    
    ' Determine width of label to use and construct formatting string for labels
    ' First, find the exponent of dblThisValue
    strWork = Format(dblThisValue, "0E+000")
    lngExponentValue = CIntSafe(Right(strWork, 4))
    
    ' Determine number of digits in dblThisValue, rounded according to lngExponentVal
    If lngExponentValue >= 0 Then
        intDigitsInLabel = 0
        strFormatString = "0"
    Else
        ' Add 1 for the decimal point
        intDigitsInLabel = -lngExponentValue + 1
        strFormatString = "0." & String(-lngExponentValue, "0")
    End If

    intDigitsInFormattedValue = Len(Format(dblThisValue, strFormatString))
    
    ConstructFormatString = strFormatString
End Function

Public Function CountOccurrenceInString(ByVal strStringToSearch As String, ByVal strSearchString As String, Optional ByVal blnCaseSensitive As Boolean = False) As Long
    ' Counts the number of times strSearchString occurs in strStringToSearch
    
    Dim lngMatchCount As Long, lngCharLoc As Long
    
    If Not blnCaseSensitive Then
        strStringToSearch = LCase(strStringToSearch)
        strSearchString = LCase(strSearchString)
    End If
    
On Error GoTo CountOccurrenceInStringErrorHandler

    If Len(strSearchString) = 0 Or Len(strStringToSearch) = 0 Then
        lngMatchCount = 0
    Else
        lngCharLoc = 1
        Do
            lngCharLoc = InStr(lngCharLoc, strStringToSearch, strSearchString)
            If lngCharLoc > 0 Then
                lngMatchCount = lngMatchCount + 1
                lngCharLoc = lngCharLoc + Len(strSearchString)
            End If
        Loop While lngCharLoc > 0
    End If
    
    CountOccurrenceInString = lngMatchCount
    Exit Function
    
CountOccurrenceInStringErrorHandler:
    Debug.Assert False
    Debug.Print "Error in CountOccurrenceInString Function: " & Err.Description
    CountOccurrenceInString = lngMatchCount
    
End Function

Public Function CountInstancesOfApp(frmThisForm As VB.Form) As Integer
    Dim lngHWnd As Long, strCaption As String, intInstances As Integer
    Dim strClassName As String, nRtn As Long, lngMaxCount As Long
    
    lngMaxCount = 100
    
    ' See if other copies of the Software are already loaded
    ' A hidden window bearing the name of the application is loaded for each instance of the app,
    '   regardless of the caption of frmMain
    '
    ' Thus, I can search for all forms with the caption frmThisForm,
    '   then examine the ClassName of the form.
    ' Windows explorer folders are class CabinetWClass
    ' All VB Apps are class ThunderMain (if in VB IDE) or ThunderRT6Main (if compiled to .Exe)
    ' In addition, if frmMain's caption is App.Title, it will also be found,
    '    However, it is class ThunderFormDC or ThunderRT6FormDC
    
    ' Finally, note that I could use App.Previnstance to see if other instances of the Application are running
    ' However, this only works for other instances with the identical version number
    ' Thus, I'll use the Window-Handle method shown below.
    
    On Error GoTo ExitCountInstances
    
    lngHWnd = frmThisForm.hwnd
    intInstances = 0
    Do
        DoEvents
        If lngHWnd = 0 Then Exit Do
        strCaption = GetWindowCaption(lngHWnd)
        
        If LCase(strCaption) = LCase(App.Title) Then
            ' Note: Usage of GetClassName from Joe Garrick, at
            '       http://www.citilink.com/~jgarrick/vbasic/tips/appact.html
            ' Must fill strClassName with spaces (or nulls) before sending to GetClassName
            ' All VB Apps are class Thunder, and each app has one window with Thunder and Main
            ' Windows explorer folders are class CabinetWClass
            
            strClassName = Space(lngMaxCount)
            nRtn = GetClassName(lngHWnd, strClassName, lngMaxCount)
            
'            If Len(strCaption) > 0 Then
'                Debug.Print "(" & lngHWnd&; ") " & strCaption & ": "; strClassName
'            End If
            
            strClassName = Trim(LCase(strClassName))
            If InStr(strClassName, "thunder") Then
                If InStr(strClassName, "main") Then
                    intInstances = intInstances + 1
                End If
            End If
        End If
        
        lngHWnd = GetNextWindow(lngHWnd, 2)
    Loop
    
ExitCountInstances:
    CountInstancesOfApp = intInstances
End Function

Public Function GetMostRecentTextBoxValue() As String
    GetMostRecentTextBoxValue = mTextBoxValueSaved
End Function

Public Function GetTemporaryDir(Optional blnFavorAPI As Boolean = True) As String
    ' Uses two different methods to get the temporary directory path
    
    Dim strTempDirViaEnviron As String, strTempDirViaAPI As String
    Dim strResult As String
    Dim lngCount As Long
    Const MAX_LENGTH = 1024
        
    ' Get temp directory using the Environ() Function
    strTempDirViaEnviron = Environ("TMP")
    If Len(strTempDirViaEnviron) = 0 Then
        strTempDirViaEnviron = Environ("TEMP")
    End If
    
    If Len(strTempDirViaEnviron) > 0 Then
        If Right(strTempDirViaEnviron, 1) <> "\" Then
            strTempDirViaEnviron = strTempDirViaEnviron & "\"
        End If
    End If
    
    ' Get temp directory using an API call
    strResult = Space(MAX_LENGTH)
    lngCount = GetTempPath(MAX_LENGTH, strResult)
    
    If lngCount > 0 Then
        If lngCount > Len(strResult) Then
            strResult = Space(lngCount + 1)
            lngCount = GetTempPath(MAX_LENGTH, strResult)
        End If
    End If
    
    If lngCount > 0 Then
        strTempDirViaAPI = Left(strResult, lngCount)
    Else
        strTempDirViaAPI = ""
    End If

    If strTempDirViaAPI = strTempDirViaEnviron Then
        GetTemporaryDir = strTempDirViaAPI
    Else
        If blnFavorAPI Then
            GetTemporaryDir = strTempDirViaAPI
        Else
            GetTemporaryDir = strTempDirViaEnviron
        End If
            
    End If
    
End Function

Private Function GetWindowCaption(lngHWnd As Long) As String
    Dim strCaption As String, lngLength As Long

    lngLength = GetWindowTextLength(lngHWnd)

    strCaption = String(lngLength, 0&)

    GetWindowText lngHWnd, strCaption, lngLength + 1
    GetWindowCaption = strCaption

End Function

Public Function CreateFolderByPath(ByVal strFolderPathToCreate As String) As Boolean
    ' strFolderPathToCreate must have a fully qualified folder path
    ' For example: C:\Temp\MyFolder
    '            or \\MyServer\MyShare\NewFolder
    ' This function will recursively step through the parent folders of the given folder,
    '  creating them as needed
    ' Thus, if strFolderPathToCreate = "C:\Temp\SubFolder\Docs\Working" and only C:\Temp exists
    '  then 3 folders will be created: SubFolder, Docs, and Working
    
    ' Returns True if the folder already exists or if it is successfully created
    ' Returns False if an error occurs or if the folder can't be created
    
    Dim fso As New FileSystemObject
    Dim strParentFolderPath As String
    Dim blnSuccess As Boolean
    
On Error GoTo CreateFolderByPathErrorHandler

    strParentFolderPath = fso.GetParentFolderName(strFolderPathToCreate)
    
    If Len(strParentFolderPath) = 0 Then Exit Function
    
    If Not fso.FolderExists(strParentFolderPath) Then
        blnSuccess = CreateFolderByPath(strParentFolderPath)
    Else
        blnSuccess = True
    End If
    
    If fso.FolderExists(strParentFolderPath) And blnSuccess Then
        If Not fso.FolderExists(strFolderPathToCreate) Then
            On Error Resume Next
            fso.CreateFolder (strFolderPathToCreate)
            If Err.Number = 0 Then
                blnSuccess = True
            Else
                Err.Clear
                blnSuccess = False
            End If
        Else
            ' Desired folder already exists
            blnSuccess = True
        End If
    Else
        blnSuccess = False
    End If
    
    CreateFolderByPath = blnSuccess
    
    Set fso = Nothing
    Exit Function

CreateFolderByPathErrorHandler:
    CreateFolderByPath = False
    
End Function

Public Function CSngSafe(strValue As String) As Single
    If IsNumeric(strValue) Then
        CSngSafe = CSng(strValue)
    End If
End Function

Public Function CStrIfNonZero(ByVal dblThisNumber As Double, Optional strAppendString As String = "", Optional intNumDecimalPlacesToRound As Integer = -1, Optional blnEmptyStringForZero As Boolean = True, Optional blnEmptyStringForNegative As Boolean = True) As String
    Dim strFormattingString As String
    
    If (blnEmptyStringForZero And dblThisNumber = 0) Then
        CStrIfNonZero = ""
    Else
        If (blnEmptyStringForNegative And dblThisNumber < 0) Then
            CStrIfNonZero = ""
        Else
            If intNumDecimalPlacesToRound = -1 Then
                CStrIfNonZero = Trim(CStr(dblThisNumber)) & strAppendString
            Else
                If intNumDecimalPlacesToRound = 0 Then
                    strFormattingString = "0"
                Else
                    strFormattingString = "0." & String(intNumDecimalPlacesToRound, "0")
                End If
                CStrIfNonZero = Format(dblThisNumber, strFormattingString) & strAppendString
            End If
        End If
    End If
End Function

' The following subroutines require the presence of the Microsoft Internet Connection Control on a form
'''Public Function DownloadWebPage(frmICControlForm as VB.Form, strHtmlLink As String, Optional boolUpdateProgressForm As Boolean = False) As String
'''    ' Returns the text of the page at strHtmlLink
'''
'''    ' This function requires a Microsoft Internet Connection control to be present on the form specified by
'''    '  frmICControlForm
'''    ' In addition, frmProgress must be present in the project
'''
'''    ' When using this function (in conjunction with the Microsoft URL Control)
'''    '   the browser appears in the web server logs as 'Microsoft URL Control - 6.00.8862'
'''
'''    If boolUpdateProgressForm Then
'''        frmProgress.UpdateCurrentSubTask "Downloading web page"
'''    End If
'''
'''    frmICControlForm.Inet1.AccessType = icUseDefault
'''    DownloadWebPage = frmICControlForm.Inet1.OpenURL(strHtmlLink, 0)
'''
'''    If boolUpdateProgressForm Then
'''        frmProgress.UpdateCurrentSubTask "Done"
'''    End If
'''End Function
'''
'''Public Function DownloadWebPageAsByteArray(frmICControlForm as VB.Form, strHtmlLink As String, Optional boolUpdateProgressForm As Boolean = False) As Variant
'''    ' Returns a byte array of the data at strHtmlLink
'''
'''    ' This function requires a Microsoft Internet Connection control to be present on the form specified by
'''    '  frmICControlForm
'''    ' In addition, frmProgress must be present in the project
'''
''''''' Note: To download a picture to the hard drive use the following code in the sub calling this function
''''''    Dim ByteArray() As Byte
''''''
''''''    ByteArray = DownloadWebPageAsByteArray(frmMain, strHtmlLink, False)
''''''
''''''    If UBound(ByteArray()) > 0 Then
''''''        ' Save the data to disk
''''''        Open strFilepath For Binary Access Write As #1
''''''        Put #1, , ByteArray()
''''''        Close #1
''''''    End If
'''''''
'''
'''    If boolUpdateProgressForm Then
'''        frmProgress.UpdateCurrentSubTask "Downloading web page"
'''    End If
'''
'''    frmICControlForm.Inet1.AccessType = icUseDefault
'''    DownloadWebPageAsByteArray = frmICControlForm.Inet1.OpenURL(strHtmlLink, icByteArray)
'''
'''    If boolUpdateProgressForm Then
'''        frmProgress.UpdateCurrentSubTask "Done"
'''    End If
'''End Function

Public Function DetermineDecimalPoint() As String
    Dim strTestNumber As String, sglConversionResult As Double
    
    ' I need to use the On Error Resume Next statement
    ' Since the Trim(Str(Cdbl(...))) statement causes an error when the
    '  user's computer is configured for using , for decimal points but not . for the
    '  thousand's separator (instead, perhaps, using a space for thousands)
    On Error Resume Next
    
    ' Determine what locale we're in (. or , for decimal point)
    strTestNumber = "5,500"
    sglConversionResult = CDbl(strTestNumber)
    If sglConversionResult = 5.5 Then
        ' Use comma as Decimal point
        DetermineDecimalPoint = ","
    Else
        ' Use period as Decimal point
        DetermineDecimalPoint = "."
    End If

End Function

Public Function FlattenStringArray(strArrayZeroBased() As String, lngArrayCount As Long, Optional strLineDelimeter As String = vbCrLf, Optional blnShowProgressFormOnLongOperation As Boolean = True, Optional blnIncludeDelimeterAfterFinalItem As Boolean = True) As String
    ' Flattens the entries in strArrayZeroBased() into a single string, separating each entry by strLineDelimeter
    ' Uses some recursive tricks to speed up this process vs. simply concatenating all the entries to a single string variable
    
    Const MIN_PROGRESS_COUNT = 2500
    
    ' Note: The following must be evenly divisible by 10
    Const CUMULATIVE_CHUNK_SIZE = 500
    
    Dim lngFillStringMaxIndex As Long
    Dim lngSrcIndex As Long
    Dim blnShowProgress As Boolean
    Dim FillStringArray() As String
    Dim FillStringCumulative As String
    
    lngFillStringMaxIndex = -1
    
    If lngArrayCount > MIN_PROGRESS_COUNT And blnShowProgressFormOnLongOperation Then blnShowProgress = True
    
    If blnShowProgress Then frmProgress.InitializeForm "Copying data to clipboard", 0, lngArrayCount, False, False, False
     
    ReDim FillStringArray(CLng(lngArrayCount / CUMULATIVE_CHUNK_SIZE) + 2)
    
    For lngSrcIndex = 0 To lngArrayCount - 1
        If lngSrcIndex Mod CUMULATIVE_CHUNK_SIZE / 10 = 0 Then
            If lngSrcIndex Mod CUMULATIVE_CHUNK_SIZE = 0 Then
                lngFillStringMaxIndex = lngFillStringMaxIndex + 1
            End If
        
            If blnShowProgress Then
                frmProgress.UpdateProgressBar lngSrcIndex
                If KeyPressAbortProcess > 1 Then Exit For
            End If
        End If
        
        FillStringArray(lngFillStringMaxIndex) = FillStringArray(lngFillStringMaxIndex) & strArrayZeroBased(lngSrcIndex) & strLineDelimeter
    
    Next lngSrcIndex
    
    If lngFillStringMaxIndex >= 0 And Not blnIncludeDelimeterAfterFinalItem Then
        FillStringArray(lngFillStringMaxIndex) = Left(FillStringArray(lngFillStringMaxIndex), Len(FillStringArray(lngFillStringMaxIndex)) - Len(strLineDelimeter))
    End If
        
    For lngSrcIndex = 0 To lngFillStringMaxIndex
        FillStringCumulative = FillStringCumulative & FillStringArray(lngSrcIndex)
    Next lngSrcIndex
    
    FlattenStringArray = FillStringCumulative
    
    If blnShowProgress Then frmProgress.HideForm
    
End Function

Public Function FolderExists(strPath As String) As Boolean
    Dim fso As New FileSystemObject
    
    If Len(strPath) > 0 Then
        
        FolderExists = fso.FolderExists(strPath)
    Else
        FolderExists = False
    End If

    Set fso = Nothing
End Function

Public Function FormatNumberAsString(ByVal dblNumber As Double, Optional lngMaxLength As Long = 10, Optional lngMaxDigitsOfPrecision As Long = 8, Optional blnUseScientificWhenTooLong As Boolean = True) As String
    
    Dim strNumberAsText As String
    Dim strZeroes As String
    
    If lngMaxDigitsOfPrecision <= 1 Then
        strZeroes = "0"
    Else
        strZeroes = "0." & String(lngMaxDigitsOfPrecision - 1, "0")
    End If
    
    dblNumber = CDbl(Format(dblNumber, strZeroes & "E+0"))

    strNumberAsText = CStr(dblNumber)
    
    If Len(strNumberAsText) > lngMaxLength Then
        If blnUseScientificWhenTooLong Then
            If lngMaxLength < 5 Then lngMaxLength = 5
            
            strZeroes = String(lngMaxLength - 5, "0")
            strNumberAsText = Format(dblNumber, "0." & strZeroes & "E+0")
        Else
            If lngMaxLength < 3 Then lngMaxLength = 3
            strNumberAsText = Round(dblNumber, lngMaxLength - 2)
        End If
    End If
    
    FormatNumberAsString = strNumberAsText
    
End Function

Public Function Factorial(Number As Integer) As Double
    ' Compute the factorial of a number; uses recursion
    ' Number should be an integer number between 0 and 170
    
    On Error GoTo FactorialOverflow
    
    If Number > 170 Then
        Debug.Print "Cannot compute factorial of a number over 170"
        Factorial = -1
        Exit Function
    End If
    
    If Number < 0 Then
        Debug.Print "Cannot compute factorial of a negative number"
        Factorial = -1
        Exit Function
    End If
    
    If Number = 0 Then
        Factorial = 1
    Else
        Factorial = Number * Factorial(Number - 1)
    End If
    
    Exit Function
    
FactorialOverflow:
    Debug.Print "Number too large"
    Factorial = -1
End Function

'''Public Function FileExists(strFilepath As String, Optional blnIncludeReadOnly As Boolean = True, Optional blnIncludeHidden As Boolean = False) As Boolean
'''    Warning: This function will match both files and folders
'''
'''    Dim strTestFile As String
'''    Dim intAttributes As Integer
'''
'''    If Len(strFilepath) = 0 Then
'''        FileExists = False
'''        Exit Function
'''    End If
'''
'''    If blnIncludeReadOnly Then intAttributes = intAttributes Or vbReadOnly
'''    If blnIncludeHidden Then intAttributes = intAttributes Or vbHidden
'''
'''    On Error Resume Next
'''    strTestFile = Dir(strFilepath, intAttributes)
'''
'''    If Len(strTestFile) > 0 Then
'''        FileExists = True
'''    Else
'''        FileExists = False
'''    End If
'''
'''End Function

Public Function FileExists(strFilePath As String) As Boolean
    Dim fso As New FileSystemObject
    
    If Len(strFilePath) > 0 Then
        FileExists = fso.FileExists(strFilePath)
    Else
        FileExists = False
    End If

    Set fso = Nothing
End Function

Public Function FileExtensionForce(ByVal strFilePath As String, ByVal strExtensionToForce As String, Optional ByVal blnReplaceExistingExtension As Boolean = True) As String
    ' Guarantees that strFilePath has the desired extension
    ' Returns strFilePath, with the extension appended if it isn't present
    '
    ' Example Call:    strFilePath = FileExtensionForce("MyTextFile", "txt")
    '                  will return MyTextFile.txt
    ' Second Example:  strFilePath = FileExtensionForce("MyTextFile.txt", "txt")
    '                  will return MyTextFile.txt
    
    Dim fso As New FileSystemObject
    Dim strExistingExtension As String
    
    strExtensionToForce = Trim(strExtensionToForce)
    If Left(strExtensionToForce, 1) = "." Then strExtensionToForce = Mid(strExtensionToForce, 2)
    
    strExistingExtension = fso.GetExtensionName(strFilePath)
    If UCase(strExistingExtension) <> UCase(strExtensionToForce) Then
        If blnReplaceExistingExtension And Len(strExistingExtension) > 0 Then
            FileExtensionForce = Left(strFilePath, Len(strFilePath) - Len(strExistingExtension)) & strExtensionToForce
        Else
            FileExtensionForce = strFilePath & "." & strExtensionToForce
        End If
    Else
        FileExtensionForce = strFilePath
    End If
    
    Set fso = Nothing
End Function

' Finds a window containg strCaption and returns the handle to the window
' Returns 0 if no matches
Public Function FindWindowCaption(ByVal strCaptionToFind As String, Optional ByRef strClassOfMatch As String) As Long
    
    Dim strCaption As String
    Dim lngHWnd As Long, nRtn As Long
    Const lngMaxCount = 32
    
    lngHWnd = GetTopWindow(0)
    Do While lngHWnd <> 0
        strCaption = GetWindowCaption(lngHWnd)
        
        If InStr(LCase(strCaption), LCase(strCaptionToFind)) Then
            nRtn = GetClassName(lngHWnd, strClassOfMatch, lngMaxCount)
            FindWindowCaption = lngHWnd
            Exit Do
        End If
        
        lngHWnd = GetNextWindow(lngHWnd, 2)
    Loop
    
End Function

Private Function GetDesktopSize(ByRef lngHeight As Long, ByRef lngWidth As Long, blnUseTwips As Boolean) As Long
    ' Determines the height and width of the desktop
    ' Returns values in Twips, if requested
    ' Returns the width
    Dim lngHWnd As Long, lngReturn As Long
    Dim lpWindowPlacement As WINDOWPLACEMENT
    
    lngHWnd = GetDesktopWindow()
    
    lngReturn = GetWindowPlacement(lngHWnd, lpWindowPlacement)
    
    With lpWindowPlacement.rcNormalPosition
        lngHeight = .Bottom - .Top
        lngWidth = .Right - .Left
    End With
    
    If blnUseTwips Then
        lngWidth = lngWidth * Screen.TwipsPerPixelX
        lngHeight = lngHeight * Screen.TwipsPerPixelY
    End If
    
    GetDesktopSize = lngWidth
End Function

Public Function GetClipboardTextSmart() As String
    Dim intAttempts As Integer, strClipboardText As String
    Const cMaxAttempts = 5
    
    intAttempts = 0
    
TryAgain:
    Err.Clear
    On Error GoTo TryAgain
    intAttempts = intAttempts + 1
    If intAttempts <= cMaxAttempts Then
        Sleep 100
        strClipboardText = Clipboard.GetText()
    Else
        strClipboardText = ""
        MsgBox Err.Description & vbCrLf & "Continuing without retrieving any clipboard text", vbInformation + vbOKOnly
    End If
    
    GetClipboardTextSmart = strClipboardText

End Function

''Public Function GetUrlSource(sURL As String, Optional boolUpdateProgressForm As Boolean = False) As String
''    ' Function from http://www.planet-source-code.com/vb/scripts/ShowCode.asp?lngWId=1&txtCodeId=24465
''
''    ' Note: When using this function, the browser appears in the web server logs as the currently installed web browser (typically MSIE 5.x or MSIE 6.x)
''
''    Dim sBuffer As String * BUFFER_LEN, iResult As Integer, sData As String
''    Dim hInternet As Long, hSession As Long, lReturn As Long
''
''    If boolUpdateProgressForm Then
''        frmProgress.UpdateCurrentSubTask "Initializing internal web browser"
''    End If
''
''    'get the handle of the current internet connection
''    hSession = InternetOpen("vb wininet", 1, vbNullString, vbNullString, 0)
''
''    'get the handle of the url
''    If hSession Then
''        hInternet = InternetOpenUrl(hSession, sURL, vbNullString, 0, IF_NO_CACHE_WRITE, 0)
''    End If
''
''    If boolUpdateProgressForm Then
''        frmProgress.UpdateCurrentSubTask "Downloading web page"
''    End If
''
''    'if we have the handle, then start reading the web page
''    If hInternet Then
''        'get the first chunk & buffer it.
''        iResult = InternetReadFile(hInternet, sBuffer, BUFFER_LEN, lReturn)
''        sData = sBuffer
''        'if there's more data then keep reading it into the buffer
''        Do While lReturn <> 0
''            If boolUpdateProgressForm Then
''                With frmProgress.lblCurrentSubTask
''                    .Caption = .Caption & "."
''                    If Len(.Caption) > 50 Then .Caption = ""
''                End With
''                DoEvents
''            End If
''
''            iResult = InternetReadFile(hInternet, sBuffer, BUFFER_LEN, lReturn)
''            sData = sData + Mid(sBuffer, 1, lReturn)
''        Loop
''    End If
''
''    'close the URL
''    iResult = InternetCloseHandle(hInternet)
''
''    If boolUpdateProgressForm Then
''        frmProgress.UpdateCurrentSubTask "Done"
''    End If
''
''    GetUrlSource = sData
''End Function

Public Function IsCharacter(TestString As String) As Boolean
    ' Returns true if the first letter of TestString is a character (i.e. a lowercase or uppercase letter)
    Dim AsciiValue As Integer
    If Len(TestString) > 0 Then
        AsciiValue = Asc(Left(TestString, 1))
        Select Case AsciiValue
        Case 65 To 90, 97 To 122
            IsCharacter = True
        Case Else
            IsCharacter = False
        End Select
    Else
        IsCharacter = False
    End If
End Function

'for use in a .bas module or class contained in
'the same project as the form(s) being tested for
Public Function IsLoaded(FormName As String) As Boolean

    Dim strFormName As String
    Dim objForm As VB.Form
    
    strFormName = UCase(FormName)
    
    For Each objForm In Forms
       If UCase(objForm.Name) = strFormName Then
         IsLoaded = True
         Exit Function
       End If
    Next

End Function

Public Function Log10(x As Double) As Double
   On Error Resume Next
   Log10 = Log(x) / Log(10#)
End Function

''''For use in a class contained in an external .dll
''Public Function IsLoaded(FormName As String, FormCollection As Object) As Boolean
''
''    Dim strFormName As String
''    Dim f as VB.Form
''
''    strFormName = ucase(FormName)
''
''    On Error Resume Next
''    For Each f In FormCollection
''       If ucase(f.Name) = strFormName Then
''         IsLoaded = True
''         Exit Function
''       End If
''    Next
''
''End Function

''
''Public Sub LaunchDefaultWebBrowser(frmParentForm as VB.Form, strHtmlAddressToView As String)
''    ' Any valid form can be passed as the parent form
''    ' Necessary since an hWnd value is required to run ShellExecute
''
''    ShellExecute frmParentForm.hWnd, "open", strHtmlAddressToView, vbNullString, vbNullString, conSwNormal
''End Sub

Public Function MatchAndSplit(ByRef strSearchString As String, strTextBefore As String, strTextAfter As String, Optional boolRemoveTextBeforeAndMatch As Boolean = True) As String
    ' Looks for strTextBefore in the string
    ' Next looks for strTextAfter in the string
    ' Returns the text between strTextBefore and strTextAfter
    ' If strTextBefore is not found, starts at the beginning of the string
    ' If strTextAfter is not found, continues to end of string
    
    ' In addition, can remove strTextBefore and the matching string from strWork, though keeping strTextAfter
    
    Dim lngMatchIndex As Long, strWork As String, strMatch As String
    
    strWork = strSearchString
    
    lngMatchIndex = InStr(strWork, strTextBefore)
    If lngMatchIndex > 0 Then
        strWork = Mid(strWork, lngMatchIndex + Len(strTextBefore))
    End If
    lngMatchIndex = InStr(strWork, strTextAfter)
    
    If lngMatchIndex > 0 Then
        strMatch = Left(strWork, lngMatchIndex - 1)
        strWork = Mid(strWork, lngMatchIndex)
    Else
        strMatch = strWork
        strWork = ""
    End If
    
    If boolRemoveTextBeforeAndMatch Then
        strSearchString = strWork
    End If
    
    MatchAndSplit = strMatch
End Function

Public Sub ParseAndSortList(strSearchList As String, ByRef strTextArrayZeroBased() As String, ByRef lngTextArrayCount As Long, Optional ByVal strDelimeters As String = ",;", Optional blnSpaceDelimeter As Boolean = True, Optional blnCarriageReturnDelimeter As Boolean = True, Optional blnSortItems As Boolean = True, Optional blnSortItemsAsNumbers As Boolean = False, Optional blnRemoveDuplicates As Boolean = True, Optional ByVal lngMaxParseCount As Long = 1000)
    ' Breaks apart strSearchList according to the delimeter specifications
    ' Places the subparts in strTextArrayZeroBased()
    ' Optionally sorts the items
    ' Optionally removes duplicate items
    '
    '
    ' Since this sub requires the use of a special, single character to represent
    '  a carriage return, if strDelimeters = "" and blnSpaceDelimeter = False, but
    '  blnCarriageReturnDelimeter = True, then we must choose a substitute single
    '  character delimeter.  We've chosen to use Chr(1).  Thus, strSearchList should
    '  not contain Chr(1) in this special case
    
    Dim strCrLfReplacement As String
    Dim lngArrayCountBeforeDuplicatesRemoval As Long
    Dim lngIndex As Long
    
    If lngMaxParseCount <= 0 Then lngMaxParseCount = 1000000000#
    
    ' Possibly add a space to strDelimeters
    If blnSpaceDelimeter Then strDelimeters = strDelimeters & " "
    If Len(strDelimeters) = 0 And blnCarriageReturnDelimeter Then
        ' Need to have at least one character in strDelimeters for the Replace() statement below to work
        ' I'll use Chr(1) since
        strDelimeters = Chr(1)
    End If
    
    If Len(strDelimeters) = 0 Then
        ' No delimeters
        ' Place all of the text in strTextArrayZeroBased(0)
        lngTextArrayCount = 1
        ReDim strTextArrayZeroBased(0)
        strTextArrayZeroBased(0) = strSearchList
        Exit Sub
    End If
    
    If blnCarriageReturnDelimeter Then
        ' To make life easier, replace all of the carriage returns in strSearchList with
        ' the first delimeter in strDelimeters (stored in strCrLfReplacement)
        strCrLfReplacement = Left(strDelimeters, 1)
        strSearchList = Replace(strSearchList, vbCrLf, strCrLfReplacement)
    End If
    
    lngTextArrayCount = ParseString(strSearchList, strTextArrayZeroBased(), lngMaxParseCount, strDelimeters, "", False, True, False)
    
    If lngTextArrayCount >= 1 Then
        ReDim Preserve strTextArrayZeroBased(0 To lngTextArrayCount - 1)
        
        If blnSortItems Then
            ShellSortString strTextArrayZeroBased(), 0, lngTextArrayCount - 1, blnSortItemsAsNumbers
        End If
        
        lngArrayCountBeforeDuplicatesRemoval = lngTextArrayCount
        If blnRemoveDuplicates Then
            RemoveDuplicates strTextArrayZeroBased(), lngTextArrayCount
        End If
    Else
        lngTextArrayCount = 0
        ReDim strTextArrayZeroBased(0)
    End If
    
End Sub

Public Function ParseString(ByVal strWork As String, ByRef strParsedVals() As String, lngParseTrackMax As Long, Optional strFieldDelimeter As String = FIELD_DELIMETER, Optional strRemaining As String, Optional boolMatchWholeDelimeter As Boolean = True, Optional boolCombineConsecutiveDelimeters As Boolean = False, Optional blnOneBaseArray As Boolean = True) As Long
    ' Scans strWork, looking for strFieldDelimeter, splitting strWork into a maximum of lngParseTrackMax parts
    '  and storing the results in strParsedVals()
    ' Note that strParsedVals() is a 1-based array if blnOneBaseArray = True (which is default)
    
    ' strFieldDelimeter may be 1 or more characters long.  If multiple characters, use
    '   boolMatchWholeDelimeter = True to treat strFieldDelimeter as just one delimeter
    ' Use boolMatchWholeDelimeter = False to treat each of the characters in strFieldDelimeter as a delimeter (token)
    ' When boolCombineConsecutiveDelimeters is true, then consecutive delimeters (like ,,, or two or more spaces) will be treated as just one delimeter
    
    ' Returns the number of values found
    ' If there was strParsedVals() gets filled to lngParseTrackMax, then the remaining text is placed in strRemaining
    
    Const DIM_CHUNK_SIZE = 10
    
    Dim lngParseTrack As Long, lngMatchIndex As Long
    Dim lngParseTrackDimCount As Long
    
    Dim lngIndexOffset As Long
    Dim lngCharLoc As Long
    
    If blnOneBaseArray Then
        lngIndexOffset = 0
    Else
        lngIndexOffset = 1
    End If
    
    lngParseTrackDimCount = DIM_CHUNK_SIZE
    
    ' Need to use On Error Resume Next here in case strParsedVals() has a fixed size (i.e. was dimmed at design time)
    On Error Resume Next
    ReDim strParsedVals(lngParseTrackDimCount + 1)      ' Must add 1 since any remainder is placed in array (following the Do While-Loop)
    On Error GoTo ParseStringErrorHandler
    
    lngParseTrack = 0
    lngMatchIndex = ParseStringFindNextDelimeter(strWork, strFieldDelimeter, boolMatchWholeDelimeter, boolCombineConsecutiveDelimeters)
    Do While lngMatchIndex > 0 And lngParseTrack < lngParseTrackMax
        lngParseTrack = lngParseTrack + 1
        If lngParseTrack >= lngParseTrackDimCount Then
            lngParseTrackDimCount = lngParseTrackDimCount + DIM_CHUNK_SIZE
            On Error Resume Next
            ReDim Preserve strParsedVals(lngParseTrackDimCount + 1)      ' Must add 1 since any remainder is placed in array (following the Do While-Loop)
            On Error GoTo ParseStringErrorHandler
        End If
        
        If lngMatchIndex > 1 Then
            strParsedVals(lngParseTrack - lngIndexOffset) = Left(strWork, lngMatchIndex - 1)
        Else
            strParsedVals(lngParseTrack - lngIndexOffset) = ""
        End If
        
        If boolMatchWholeDelimeter Then
            strWork = Mid(strWork, lngMatchIndex + Len(strFieldDelimeter))
        Else
            strWork = Mid(strWork, lngMatchIndex + 1)
            If boolCombineConsecutiveDelimeters Then
                ' Need to check for, and remove, any delimeters at the end of strParsedVals(lngParseTrack - lngIndexOffset)
                Do
                    lngCharLoc = ParseStringFindNextDelimeter(strParsedVals(lngParseTrack - lngIndexOffset), strFieldDelimeter, False, True)
                    If lngCharLoc > 0 Then
                        strParsedVals(lngParseTrack - lngIndexOffset) = Left(strParsedVals(lngParseTrack - lngIndexOffset), lngCharLoc - 1) & Mid(strParsedVals(lngParseTrack - lngIndexOffset), lngCharLoc + 1)
                    End If
                Loop While lngCharLoc > 0
            End If
        End If
        lngMatchIndex = ParseStringFindNextDelimeter(strWork, strFieldDelimeter, boolMatchWholeDelimeter, boolCombineConsecutiveDelimeters)
    Loop
    
    If Len(strWork) > 0 Then
        ' Items still remain; append to strParsedVals() or place in strRemaining
        If lngParseTrack < lngParseTrackMax Then
            lngParseTrack = lngParseTrack + 1
            strParsedVals(lngParseTrack - lngIndexOffset) = strWork
        Else
            strRemaining = strWork
        End If
    End If
    
    ParseString = lngParseTrack
    Exit Function

ParseStringErrorHandler:
    Debug.Assert False
    Debug.Print "Error with ParseString: " & Err.Description
    ParseString = lngParseTrack
    
End Function

Public Function ParseStringValues(ByVal strWork As String, ByRef intParsedVals() As Integer, intParseTrackMax As Integer, Optional strFieldDelimeter As String = FIELD_DELIMETER, Optional strRemaining As String, Optional boolMatchWholeDelimeter As Boolean = True, Optional boolCombineConsecutiveDelimeters As Boolean = False, Optional blnOneBaseArray As Boolean = True) As Integer
    ' See ParseString for parameter descriptions
    
    Dim intParseTrack As Integer
    Dim intIndex As Integer, intMaxIndexToCopy As Integer
    Dim strParsedVals() As String
    
    If intParseTrackMax < 0 Then intParseTrackMax = 0
    ReDim strParsedVals(intParseTrackMax + 1)
    
    ' Need to use On Error Resume Next here in case intParsedVals() has a fixed size (i.e. was dimmed at design time)
    On Error Resume Next
    ReDim intParsedVals(intParseTrackMax + 1)
    On Error GoTo ParseStringValuesErrorHandler
    
    intParseTrack = ParseString(strWork, strParsedVals(), CLng(intParseTrackMax), strFieldDelimeter, strRemaining, boolMatchWholeDelimeter, boolCombineConsecutiveDelimeters, blnOneBaseArray)
    
    intMaxIndexToCopy = intParseTrackMax
    If UBound(intParsedVals) < intMaxIndexToCopy Then
        intMaxIndexToCopy = UBound(intParsedVals)
    End If
    
    For intIndex = 0 To intMaxIndexToCopy
        If IsNumeric(strParsedVals(intIndex)) Then
            intParsedVals(intIndex) = CInt(strParsedVals(intIndex))
        Else
            intParsedVals(intIndex) = 0
        End If
    Next intIndex
    
    ParseStringValues = intParseTrack
    Exit Function

ParseStringValuesErrorHandler:
    Debug.Assert False
    Debug.Print "Error with ParseStringValues: " & Err.Description
    ParseStringValues = intParseTrack

End Function

Public Function ParseStringValuesDbl(ByVal strWork As String, ByRef dblParsedVals() As Double, intParseTrackMax As Integer, Optional strFieldDelimeter As String = FIELD_DELIMETER, Optional strRemaining As String, Optional boolMatchWholeDelimeter As Boolean = True, Optional boolCombineConsecutiveDelimeters As Boolean = False, Optional blnOneBaseArray As Boolean = True) As Integer
    ' See ParseStringText
    
    Dim intParseTrack As Integer
    Dim intIndex As Integer, intMaxIndexToCopy As Integer
    Dim strParsedVals() As String
    
    If intParseTrackMax < 0 Then intParseTrackMax = 0
    ReDim strParsedVals(intParseTrackMax + 1)
    
    ' Need to use On Error Resume Next here in case dblParsedVals() has a fixed size (i.e. was dimmed at design time)
    On Error Resume Next
    ReDim dblParsedVals(intParseTrackMax + 1)
    On Error GoTo ParseStringValuesDblErrorHandler
    
    intParseTrack = ParseString(strWork, strParsedVals(), CLng(intParseTrackMax), strFieldDelimeter, strRemaining, boolMatchWholeDelimeter, boolCombineConsecutiveDelimeters, blnOneBaseArray)
    
    intMaxIndexToCopy = intParseTrackMax
    If UBound(dblParsedVals) < intMaxIndexToCopy Then
        intMaxIndexToCopy = UBound(dblParsedVals)
    End If
    
    For intIndex = 0 To intMaxIndexToCopy
        If IsNumeric(strParsedVals(intIndex)) Then
            dblParsedVals(intIndex) = CDbl(strParsedVals(intIndex))
        Else
            dblParsedVals(intIndex) = 0
        End If
    Next intIndex
    
    ParseStringValuesDbl = intParseTrack
    Exit Function
    
ParseStringValuesDblErrorHandler:
    Debug.Assert False
    Debug.Print "Error with ParseStringValuesDbl: " & Err.Description
    ParseStringValuesDbl = intParseTrack
    
End Function
 
Public Function ParseStringFindCrlfIndex(ByRef strWork As String, ByRef intDelimeterLength As Integer) As Long
    ' First looks for vbCrLf in strWork
    ' Returns index if found, setting intDelimeterLength to 2
    ' If not found, uses ParseStringFindNextDelimeter to search for just CR or just LF,
    '  returning location and setting intDelimeterLength to 1
    
    Dim lngCrLfLoc As Long
    
    lngCrLfLoc = InStr(strWork, vbCrLf)
    If lngCrLfLoc = 0 Then
        ' CrLf not found; look for just Cr or just LF
        lngCrLfLoc = ParseStringFindNextDelimeter(strWork, vbCrLf, False)
        intDelimeterLength = 1
    Else
        intDelimeterLength = 2
    End If

    ParseStringFindCrlfIndex = lngCrLfLoc
End Function
 
Private Function ParseStringFindNextDelimeter(ByVal strWork As String, strFieldDelimeter As String, Optional boolMatchWholeDelimeter As Boolean = True, Optional boolCombineConsecutiveDelimeters As Boolean = False) As Long
    ' Scans strWork, looking for next delimeter (token)
    ' strFieldDelimeter may be 1 or more characters long.  If multiple characters, use
    '   boolMatchWholeDelimeter = True to treat strFieldDelimeter as just one delimeter
    ' Use boolMatchWholeDelimeter = False to treat each of the characters in strFieldDelimeter as a delimeter (token)
    
    Dim intFieldDelimeterLength As Integer, intDelimeterIndex As Integer
    Dim lngMatchIndex As Long, lngSmallestMatchIndex As Long
    Dim blnDelimeterMatched As Boolean
    
    intFieldDelimeterLength = Len(strFieldDelimeter)
    
    If intFieldDelimeterLength = 0 Then
        lngSmallestMatchIndex = 0
    Else
        If boolMatchWholeDelimeter Or intFieldDelimeterLength = 1 Then
            lngSmallestMatchIndex = InStr(strWork, strFieldDelimeter)
        Else
            ' Look for each of the characters in strFieldDelimeter, returning the smallest nonzero index found
            lngSmallestMatchIndex = 0
            For intDelimeterIndex = 1 To Len(strFieldDelimeter)
                lngMatchIndex = InStr(strWork, Mid(strFieldDelimeter, intDelimeterIndex, 1))
                If lngMatchIndex > 0 Then
                    If lngMatchIndex < lngSmallestMatchIndex Or lngSmallestMatchIndex = 0 Then
                        lngSmallestMatchIndex = lngMatchIndex
                    End If
                End If
            Next intDelimeterIndex
        End If
    
        ' If boolCombineConsecutiveDelimeters is true, then examine adjacent text for more delimeters, returning location of final delimeter
        If boolCombineConsecutiveDelimeters Then
            lngMatchIndex = lngSmallestMatchIndex + 1
            Do While lngMatchIndex <= Len(strWork)
                If boolMatchWholeDelimeter Or intFieldDelimeterLength = 1 Then
                    If Mid(strWork, lngMatchIndex, intFieldDelimeterLength) = strFieldDelimeter Then
                        lngMatchIndex = lngMatchIndex + intFieldDelimeterLength
                    Else
                        Exit Do
                    End If
                Else
                    blnDelimeterMatched = False
                    For intDelimeterIndex = 1 To Len(strFieldDelimeter)
                        If Mid(strWork, lngMatchIndex, 1) = Mid(strFieldDelimeter, intDelimeterIndex, 1) Then
                            blnDelimeterMatched = True
                            Exit For
                        End If
                    Next intDelimeterIndex
                    If blnDelimeterMatched Then
                        lngMatchIndex = lngMatchIndex + 1
                    Else
                        Exit Do
                    End If
                End If
            Loop
            lngSmallestMatchIndex = lngMatchIndex - 1
        End If
    
    End If
    
    ParseStringFindNextDelimeter = lngSmallestMatchIndex

End Function

Public Sub RemoveDuplicates(ByRef strTextArrayZeroBased() As String, ByRef lngArrayCount As Long)
    Dim lngIndex As Long
    Dim lngCompareIndex As Long
    Dim lngShiftIndex As Long
    
    lngIndex = 0
    Do While lngIndex < lngArrayCount
        lngCompareIndex = lngArrayCount - 1
        Do While lngCompareIndex > lngIndex
            If strTextArrayZeroBased(lngIndex) = strTextArrayZeroBased(lngCompareIndex) Then
                ' Remove duplicate item
                For lngShiftIndex = lngCompareIndex To lngArrayCount - 2
                    strTextArrayZeroBased(lngShiftIndex) = strTextArrayZeroBased(lngShiftIndex + 1)
                Next lngShiftIndex
                lngArrayCount = lngArrayCount - 1
                strTextArrayZeroBased(lngArrayCount) = ""
                If lngCompareIndex >= lngArrayCount Then Exit Do
            Else
                lngCompareIndex = lngCompareIndex - 1
            End If
        Loop
        lngIndex = lngIndex + 1
    Loop
    
End Sub

Public Function ReplaceSubString(ByRef strSearchString As String, strTextToFind As String, strTextToReplaceWith As String) As Boolean
    ' Returns true if a change was made to strSearchString
    
    Dim lngMatchIndex As Long, intSearchTextLength As Integer, boolReplaced As Boolean
    
    boolReplaced = False
    intSearchTextLength = Len(strTextToFind)
    
    Do
        lngMatchIndex = InStr(strSearchString, strTextToFind)
        If lngMatchIndex > 0 Then
            strSearchString = Left(strSearchString, lngMatchIndex - 1) & strTextToReplaceWith & Mid(strSearchString, lngMatchIndex + intSearchTextLength)
            boolReplaced = True
        End If
    Loop While lngMatchIndex > 0
    
    ReplaceSubString = boolReplaced
    
End Function

Public Function RoundToNearest(ByVal dblNumberToRound As Double, ByVal lngMultipleToRoundTo As Long, ByVal blnRoundUp As Boolean) As Long
    ' Rounds a number to the nearest Multiple specified
    ' If blnRoundUp = True, then always rounds up
    ' If blnRoundUp = False, then always rounds down
    Dim lngRoundedNumber As Long
    
    If lngMultipleToRoundTo = 0 Then lngMultipleToRoundTo = 1
    
    ' Use Int() to get the floor of the number
    lngRoundedNumber = Int(dblNumberToRound / CDbl(lngMultipleToRoundTo)) * lngMultipleToRoundTo

    If blnRoundUp And lngRoundedNumber < dblNumberToRound Then
        lngRoundedNumber = lngRoundedNumber + lngMultipleToRoundTo
    End If
    RoundToNearest = lngRoundedNumber
End Function

Public Function SelectFile(ByVal Ownerhwnd As Long, ByVal sTitle As String, Optional ByRef strStartPath As String = "", Optional blnSaveFile As Boolean = False, Optional strDefaultFileName As String = "", Optional ByVal strFileFilterCodes As String = "All Files (*.*)|*.*|Text Files (*.txt)|*.txt", Optional ByRef intFilterIndexDefault As Integer = 1, Optional blnFileMustExistOnOpen As Boolean = True) As String
    ' Returns file name if user selects a file (for opening) or enters a valid path (for saving)
    ' Returns "" if user canceled dialog
    ' If strStartPath = "", then uses last directory for default directory
    ' Updates strStartPath to contain the directory of the selected file
    ' Updates intFilterIndexDefault to have the filter index of the selected file (if a valid file is chosen)
    
    Dim ofDlg As OPENFILENAME
    Dim Res As Long
    Dim Chr0Pos As Integer
    
    Dim sFilter As String
    Dim nFilterInd As Integer
    Dim strSelectedFile As String
    Dim strSuggestedFileName As String
    
    If Len(strFileFilterCodes) = 0 Then
        sFilter = "All Files" & Chr(0) & "*.*"
        nFilterInd = 1
    Else
        sFilter = Replace(strFileFilterCodes, "|", Chr(0))
        sFilter = sFilter & Chr(0)
        nFilterInd = intFilterIndexDefault
    End If
    
    If Len(strStartPath) > 0 Then
        mCurrentDirectory = strStartPath & vbNullChar
    End If
    
    If Len(mCurrentDirectory) = 0 Then mCurrentDirectory = vbNullChar
    
    strSuggestedFileName = strDefaultFileName
    
    ' GetSaveFileName() doesn't like having colons in the suggested filename
    strSuggestedFileName = Replace(strSuggestedFileName, ":", "")
    
    With ofDlg
        .lStructSize = Len(ofDlg)
        .hwndOwner = Ownerhwnd
        .hInstance = App.hInstance
        .lpstrFilter = sFilter
        .nFilterIndex = nFilterInd
        If Len(strSuggestedFileName) > 0 Then
            .lpstrFile = strSuggestedFileName & String(MAX_PATH, 0)
        Else
            .lpstrFile = String(MAX_PATH, 0)
        End If
        .nMaxFile = Len(.lpstrFile) - 1
        .lpstrFileTitle = .lpstrFile
        .nMaxFileTitle = .nMaxFile
        .lpstrInitialDir = mCurrentDirectory
        .lpstrTitle = sTitle
        If blnSaveFile Then
            .flags = OFN_LONGNAMES Or OFN_OVERWRITEPROMPT Or OFN_HIDEREADONLY
        Else
            If blnFileMustExistOnOpen Then
                .flags = OFN_FILEMUSTEXIST
            Else
                .flags = 0
            End If
        End If
    End With
    
    If blnSaveFile Then
        Res = GetSaveFileName(ofDlg)
    Else
        Res = GetOpenFileName(ofDlg)
    End If
    
    ' Give Windows a chance to refresh (i.e., close the dialog)
    DoEvents

    If Res = 0 Then
        SelectFile = ""
    Else
        Chr0Pos = InStr(1, ofDlg.lpstrFile, Chr(0))
        If Chr0Pos > 0 Then
            strSelectedFile = Left(ofDlg.lpstrFile, Chr0Pos - 1)
        Else
            strSelectedFile = Trim(ofDlg.lpstrFile)
        End If
        
        StripFullPath strSelectedFile, mCurrentDirectory
        strStartPath = mCurrentDirectory
        intFilterIndexDefault = ofDlg.nFilterIndex
    
        SelectFile = strSelectedFile
    End If
    
End Function

Private Function SelectValue(ByRef sngInputArrayOneBased() As Single, lngArrayCount As Long, lngElementToSelect As Long) As Single
    ' Rearranges sngInputArrayZeroBased such that the lngElementToSelect'th element is in sngInputArrayZeroBased(lngElementToSelect)
    ' Code is from Numerical Recipes In C (Section 8.5, page 375)
    
    ' Returns the kth smallest value in the array arr[1...n].  The input array will be rearranged
    ' to have this value in location arr[k], with all smaller elements moved to arr[1..k-1] (in
    ' arbitrary order) and all larger elements in arr[k+1..n] (also in arbitrary order).

    Dim i As Long, ir As Long, j As Long, l As Long, lngMidPoint As Long
    
    Dim a As Single

    l = 1
    ir = lngArrayCount
    
    ' Loop until the Exit Do statement is reached
    Do
        
        If ir <= l + 1 Then
            If ir = l + 1 And sngInputArrayOneBased(ir) < sngInputArrayOneBased(l) Then
                SwapSingle sngInputArrayOneBased(l), sngInputArrayOneBased(ir)
            End If
            Exit Do
        Else
            lngMidPoint = Int((l + ir) / 2)
            SwapSingle sngInputArrayOneBased(lngMidPoint), sngInputArrayOneBased(l + 1)
            
            If sngInputArrayOneBased(l) > sngInputArrayOneBased(ir) Then
                SwapSingle sngInputArrayOneBased(l), sngInputArrayOneBased(ir)
            End If
            If sngInputArrayOneBased(l + 1) > sngInputArrayOneBased(ir) Then
                SwapSingle sngInputArrayOneBased(l + 1), sngInputArrayOneBased(ir)
            End If
            If sngInputArrayOneBased(l) > sngInputArrayOneBased(l + 1) Then
                SwapSingle sngInputArrayOneBased(l), sngInputArrayOneBased(l + 1)
            End If
            i = l + 1
            j = ir
            a = sngInputArrayOneBased(l + 1)
            
            Do
                Do
                    i = i + 1
                Loop While sngInputArrayOneBased(i) < a
                
                Do
                    j = j - 1
                Loop While sngInputArrayOneBased(j) > a
                
                If j < i Then Exit Do
                
                SwapSingle sngInputArrayOneBased(i), sngInputArrayOneBased(j)
            Loop
            sngInputArrayOneBased(l + 1) = sngInputArrayOneBased(j)
            sngInputArrayOneBased(j) = a
            If j >= lngElementToSelect Then ir = j - 1
            If j <= lngElementToSelect Then l = i
            
        End If
    Loop
    
    SelectValue = sngInputArrayOneBased(lngElementToSelect)
    
''    Debug.Print "Checking median; should be " & sngInputArrayOneBased(lngElementToSelect)
''    For j = 1 To lngArrayCount
''        Debug.Print sngInputArrayOneBased(j)
''    Next j
''    Debug.Print ""
    
End Function

Public Function SetCheckBox(ByRef chkThisCheckBox As VB.CheckBox, blnIsChecked As Boolean)
    If blnIsChecked Then
        chkThisCheckBox.value = vbChecked
    Else
        chkThisCheckBox.value = vbUnchecked
    End If
End Function

Public Sub SetMostRecentTextBoxValue(strNewText As String)
    mTextBoxValueSaved = strNewText
End Sub

Public Sub ShellSortLong(ByRef lngArray() As Long, ByVal lngLowIndex As Long, ByVal lngHighIndex As Long)
    Dim lngCount As Long
    Dim lngIncrement As Long
    Dim lngIndex As Long
    Dim lngIndexCompare As Long
    Dim lngCompareVal As Long

On Error GoTo ShellSortLongErrorHandler

' sort array[lngLowIndex..lngHighIndex]

    ' compute largest increment
    lngCount = lngHighIndex - lngLowIndex + 1
    lngIncrement = 1
    If (lngCount < 14) Then
        lngIncrement = 1
    Else
        Do While lngIncrement < lngCount
            lngIncrement = 3 * lngIncrement + 1
        Loop
        lngIncrement = lngIncrement \ 3
        lngIncrement = lngIncrement \ 3
    End If

    Do While lngIncrement > 0
        ' sort by insertion in increments of lngIncrement
        For lngIndex = lngLowIndex + lngIncrement To lngHighIndex
            lngCompareVal = lngArray(lngIndex)
            For lngIndexCompare = lngIndex - lngIncrement To lngLowIndex Step -lngIncrement
                ' Use <= to sort ascending; Use > to sort descending
                If lngArray(lngIndexCompare) <= lngCompareVal Then Exit For
                lngArray(lngIndexCompare + lngIncrement) = lngArray(lngIndexCompare)
            Next lngIndexCompare
            lngArray(lngIndexCompare + lngIncrement) = lngCompareVal
        Next lngIndex
        lngIncrement = lngIncrement \ 3
    Loop
    
    Exit Sub

ShellSortLongErrorHandler:
    Debug.Assert False
End Sub

' Shell Sort
Public Sub ShellSortSingle(ByRef sngArray() As Single, ByVal lngLowIndex As Long, ByVal lngHighIndex As Long)
    Dim lngCount As Long
    Dim lngIncrement As Long
    Dim lngIndex As Long
    Dim lngIndexCompare As Long
    Dim sngCompareVal As Single

On Error GoTo ShellSortSingleErrorHandler

' sort array[lngLowIndex..lngHighIndex]

    ' compute largest increment
    lngCount = lngHighIndex - lngLowIndex + 1
    lngIncrement = 1
    If (lngCount < 14) Then
        lngIncrement = 1
    Else
        Do While lngIncrement < lngCount
            lngIncrement = 3 * lngIncrement + 1
        Loop
        lngIncrement = lngIncrement \ 3
        lngIncrement = lngIncrement \ 3
    End If

    Do While lngIncrement > 0
        ' sort by insertion in increments of lngIncrement
        For lngIndex = lngLowIndex + lngIncrement To lngHighIndex
            sngCompareVal = sngArray(lngIndex)
            For lngIndexCompare = lngIndex - lngIncrement To lngLowIndex Step -lngIncrement
                ' Use <= to sort ascending; Use > to sort descending
                If sngArray(lngIndexCompare) <= sngCompareVal Then Exit For
                sngArray(lngIndexCompare + lngIncrement) = sngArray(lngIndexCompare)
            Next lngIndexCompare
            sngArray(lngIndexCompare + lngIncrement) = sngCompareVal
        Next lngIndex
        lngIncrement = lngIncrement \ 3
    Loop

''    Debug.Assert VerifySort(sngArray(), lngLowIndex, lngHighIndex)
    Exit Sub

ShellSortSingleErrorHandler:
    Debug.Assert False
End Sub


' This sub is about 3 times slower than the above ShellSortSingle routine
Public Sub ShellSortSingleOld(ByRef sngArray() As Single, ByVal lngLowIndex As Long, ByVal lngHighIndex As Long)
    Dim sngSwap As Single

    ' Sort the list via a shell sort
    Dim lngMaxRow As Long, lngOffSet As Long, lngLimit As Long, lngSwitch As Long
    Dim lngRow As Long

    ' Set comparison lngOffset to half the number of records
    lngMaxRow = lngHighIndex
    lngOffSet = lngMaxRow \ 2

    ' When just two data points in the array, need to make sure at least one comparison occurs
    If (lngHighIndex - lngLowIndex - 1) = 2 Then lngOffSet = 1

    Do While lngOffSet > 0          ' Loop until lngOffset gets to zero.

        lngLimit = lngMaxRow - lngOffSet
        Do
            lngSwitch = 0         ' Assume no switches at this lngOffset.

            ' Compare elements and Switch ones out of order:
            For lngRow = lngLowIndex To lngLimit
                ' Use > to sort ascending, < to sort descending
                If sngArray(lngRow) > sngArray(lngRow + lngOffSet) Then
                    sngSwap = sngArray(lngRow + lngOffSet)
                    sngArray(lngRow + lngOffSet) = sngArray(lngRow)
                    sngArray(lngRow) = sngSwap
                    lngSwitch = lngRow
                End If
            Next lngRow

            ' Sort on next pass only to where last lngSwitch was made:
            lngLimit = lngSwitch - lngOffSet
        Loop While lngSwitch

        ' No switches at last lngOffset, try one half as big:
        lngOffSet = lngOffSet \ 2

    Loop
End Sub

Public Sub ShellSortString(ByRef strArray() As String, ByVal lngLowIndex As Long, ByVal lngHighIndex As Long, Optional blnSortItemsAsNumbers As Boolean = False)
    Dim lngCount As Long
    Dim lngIncrement As Long
    Dim lngIndex As Long
    Dim lngIndexCompare As Long
    Dim strCompareVal As String

On Error GoTo ShellSortStringErrorHandler

    ' sort array[lngLowIndex..lngHighIndex]

    ' compute largest increment
    lngCount = lngHighIndex - lngLowIndex + 1
    lngIncrement = 1
    If (lngCount < 14) Then
        lngIncrement = 1
    Else
        Do While lngIncrement < lngCount
            lngIncrement = 3 * lngIncrement + 1
        Loop
        lngIncrement = lngIncrement \ 3
        lngIncrement = lngIncrement \ 3
    End If

    If blnSortItemsAsNumbers Then
        Do While lngIncrement > 0
            ' sort by insertion in increments of lngIncrement
            For lngIndex = lngLowIndex + lngIncrement To lngHighIndex
                strCompareVal = strArray(lngIndex)
                For lngIndexCompare = lngIndex - lngIncrement To lngLowIndex Step -lngIncrement
                    ' Use <= to sort ascending; Use > to sort descending
                    If Val(strArray(lngIndexCompare)) <= Val(strCompareVal) Then Exit For
                    strArray(lngIndexCompare + lngIncrement) = strArray(lngIndexCompare)
                Next lngIndexCompare
                strArray(lngIndexCompare + lngIncrement) = strCompareVal
            Next lngIndex
            lngIncrement = lngIncrement \ 3
        Loop
    Else
        Do While lngIncrement > 0
            ' sort by insertion in increments of lngIncrement
            For lngIndex = lngLowIndex + lngIncrement To lngHighIndex
                strCompareVal = strArray(lngIndex)
                For lngIndexCompare = lngIndex - lngIncrement To lngLowIndex Step -lngIncrement
                    ' Use <= to sort ascending; Use > to sort descending
                    If strArray(lngIndexCompare) <= strCompareVal Then Exit For
                    strArray(lngIndexCompare + lngIncrement) = strArray(lngIndexCompare)
                Next lngIndexCompare
                strArray(lngIndexCompare + lngIncrement) = strCompareVal
            Next lngIndex
            lngIncrement = lngIncrement \ 3
        Loop
    End If
    
    Exit Sub
    
ShellSortStringErrorHandler:
Debug.Assert False

End Sub

' Quick sort
Public Sub QuickSort(ByRef sngArray() As Single, ByVal lngLowIndex As Long, ByVal lngHighIndex As Long)
    ' From http://oopweb.com/Algorithms/Documents/Sman/Volume/s_vsq2.txt
    Dim lbStack(32) As Long
    Dim ubStack(32) As Long
    Dim lngStackPointer As Long              ' stack pointer
    Dim lngCurrentLowerBoundIndex As Long             ' current lower-bound
    Dim lngCurrentUpperBoundIndex As Long             ' current upper-bound
    Dim lngPivotIndex As Long               ' index to pivot
    Dim i As Long
    Dim j As Long
    Dim m As Long
    Dim sngSwapVal As Single            ' temp used for exchanges

    lbStack(0) = lngLowIndex
    ubStack(0) = lngHighIndex
    lngStackPointer = 0
    Do While lngStackPointer >= 0
        lngCurrentLowerBoundIndex = lbStack(lngStackPointer)
        lngCurrentUpperBoundIndex = ubStack(lngStackPointer)

        Do While (lngCurrentLowerBoundIndex < lngCurrentUpperBoundIndex)

            ' select pivot and exchange with 1st element
            lngPivotIndex = lngCurrentLowerBoundIndex + (lngCurrentUpperBoundIndex - lngCurrentLowerBoundIndex) \ 2

            ' exchange lngCurrentLowerBoundIndex, lngPivotIndex
            sngSwapVal = sngArray(lngCurrentLowerBoundIndex)
            sngArray(lngCurrentLowerBoundIndex) = sngArray(lngPivotIndex)
            sngArray(lngPivotIndex) = sngSwapVal

            ' partition into two segments
            i = lngCurrentLowerBoundIndex + 1
            j = lngCurrentUpperBoundIndex
            Do
                Do While i < j
                    If sngArray(lngCurrentLowerBoundIndex) <= sngArray(i) Then Exit Do
                    i = i + 1
                Loop

                Do While j >= i
                    If sngArray(j) <= sngArray(lngCurrentLowerBoundIndex) Then Exit Do
                    j = j - 1
                Loop

                If i >= j Then Exit Do

                ' exchange i, j
                sngSwapVal = sngArray(i)
                sngArray(i) = sngArray(j)
                sngArray(j) = sngSwapVal

                j = j - 1
                i = i + 1
            Loop

            ' pivot belongs in sngArray[j]
            ' exchange lngCurrentLowerBoundIndex, j
            sngSwapVal = sngArray(lngCurrentLowerBoundIndex)
            sngArray(lngCurrentLowerBoundIndex) = sngArray(j)
            sngArray(j) = sngSwapVal

            m = j

            ' keep processing smallest segment, and stack largest
            If m - lngCurrentLowerBoundIndex <= lngCurrentUpperBoundIndex - m Then
                If m + 1 < lngCurrentUpperBoundIndex Then
                    lbStack(lngStackPointer) = m + 1
                    ubStack(lngStackPointer) = lngCurrentUpperBoundIndex
                    lngStackPointer = lngStackPointer + 1
                End If
                lngCurrentUpperBoundIndex = m - 1
            Else
                If m - 1 > lngCurrentLowerBoundIndex Then
                    lbStack(lngStackPointer) = lngCurrentLowerBoundIndex
                    ubStack(lngStackPointer) = m - 1
                    lngStackPointer = lngStackPointer + 1
                End If
                lngCurrentLowerBoundIndex = m + 1
            End If
        Loop
        lngStackPointer = lngStackPointer - 1
    Loop
    
    Debug.Assert VerifySort(sngArray(), lngLowIndex, lngHighIndex)
End Sub

Public Sub TestSortRoutines()
    Const DATA_COUNT = 10000
    Const LOOP_COUNT = 25
    Const MOD_CHUNK = 1
    Dim sngTest(DATA_COUNT) As Single, sngArrayToSort(DATA_COUNT) As Single
    Dim lngRandomNumberSeed As Long
    Dim lngIndex As Long, lngLoop As Long
    Dim lngStartTime As Long, lngStopTime As Long
    
    Debug.Print "Testing Data_Count = " & DATA_COUNT
    lngRandomNumberSeed = Timer
    
    ' Call Rnd() with a negative number before calling Randomize() lngRandomNumberSeed in order to
    '  guarantee that we get the same order of random numbers each time
    Call Rnd(-1)
    Randomize lngRandomNumberSeed
    lngStartTime = GetTickCount()
    For lngLoop = 0 To LOOP_COUNT - 1
        If lngLoop Mod MOD_CHUNK = 0 Then Debug.Print ".";
        
        For lngIndex = 0 To DATA_COUNT - 1
            sngTest(lngIndex) = Rnd(1)
            sngArrayToSort(lngIndex) = sngTest(lngIndex)
        Next lngIndex
        
        ' Call ShellSort
        ShellSortSingleOld sngArrayToSort(), 0, DATA_COUNT - 1
        
        ' Restore sngArrayToSort() to original order
        For lngIndex = 0 To DATA_COUNT - 1
            sngArrayToSort(lngIndex) = sngTest(lngIndex)
        Next lngIndex
    Next lngLoop
    lngStopTime = GetTickCount()
    Debug.Print "ShellSortOld took " & lngStopTime - lngStartTime & " msec total"
    
    
    Call Rnd(-1)
    Randomize lngRandomNumberSeed
    lngStartTime = GetTickCount()
    For lngLoop = 0 To LOOP_COUNT - 1
        If lngLoop Mod MOD_CHUNK = 0 Then Debug.Print ".";
        
        For lngIndex = 0 To DATA_COUNT - 1
            sngTest(lngIndex) = Rnd(1)
            sngArrayToSort(lngIndex) = sngTest(lngIndex)
        Next lngIndex
        
        ' Call ShellSort
        ShellSortSingle sngArrayToSort(), 0, DATA_COUNT - 1
        
        ' Restore sngArrayToSort() to original order
        For lngIndex = 0 To DATA_COUNT - 1
            sngArrayToSort(lngIndex) = sngTest(lngIndex)
        Next lngIndex
    Next lngLoop
    lngStopTime = GetTickCount()
    Debug.Print "ShellSort took " & lngStopTime - lngStartTime & " msec total"
    
    
    Call Rnd(-1)
    Randomize lngRandomNumberSeed
    lngStartTime = GetTickCount()
    For lngLoop = 0 To LOOP_COUNT - 1
        If lngLoop Mod MOD_CHUNK = 0 Then Debug.Print ".";
        
        For lngIndex = 0 To DATA_COUNT - 1
            sngTest(lngIndex) = Rnd(1)
            sngArrayToSort(lngIndex) = sngTest(lngIndex)
        Next lngIndex
        
        ' Call QuickSort
        QuickSort sngArrayToSort(), 0, DATA_COUNT - 1
        
        ' Restore sngArrayToSort() to original order
        For lngIndex = 0 To DATA_COUNT - 1
            sngArrayToSort(lngIndex) = sngTest(lngIndex)
        Next lngIndex
    Next lngLoop
    lngStopTime = GetTickCount()
    Debug.Print "QuickSort took " & lngStopTime - lngStartTime & " msec total"
        
End Sub

Public Sub SizeAndCenterWindow(frmThisForm As VB.Form, Optional intCenterMode As Integer = 0, Optional lngWindowWidth As Long = -1, Optional lngWindowHeight As Long = -1, Optional boolSizeAndCenterOnlyOncePerProgramSession As Boolean = True, Optional intDualMonitorToUse As Integer = -1)
    ' Sub revision 1.2
    
    ' Center Mode uses one of the following:
    '    Public Const cWindowExactCenter = 0
    '    Public Const cWindowUpperThird = 1
    '    Public Const cWindowLowerThird = 2
    '    Public Const cWindowMiddleLeft = 3
    '    Public Const cWindowMiddleRight = 4
    '    Public Const cWindowTopCenter = 5
    '    Public Const cWindowBottomCenter = 6
    '    Public Const cWindowBottomRight = 7
    '    Public Const cWindowBottomLeft = 8
    '    Public Const cWindowTopRight = 9
    '    Public Const cWindowTopLeft = 10
    
    ' This sub routine properly recognizes dual monitors, centering the form to just one monitor
    
    ' lngWindowWidth and lngWindowHeight are in twips (there are 15 twips in one pixel)
    ' intDualMonitorToUse can be 0 or 1, signifying the first or second monitor
    ' boolSizeAndCenterOnlyOncePerProgramSession is useful when the SizeAndCenterWindow sub is called from the Form_Activate sub of a form
    '  Note: It is suggested that this be set to false if called from Form_Load in case the user closes the form (thus unloading it)
    
    Const MAX_RESIZE_FORMS_TO_REMEMBER = 100
    Dim lngWindowAreaWidth As Long, lngWindowAreaHeight As Long, dblAspectRatio As Double
    Dim lngWorkingAreaWidth As Long, lngWorkingAreaHeight As Long
    Dim boolDualMonitor As Boolean, boolHorizontalDual As Boolean
    Dim lngWindowTopToSet As Long, lngWindowLeftToSet As Long
    Dim frmMainAppForm As VB.Form
    Static strFormsCentered(MAX_RESIZE_FORMS_TO_REMEMBER) As String     ' 0-based array
    Static intFormsCenteredCount As Integer
    Dim boolSubCalledPreviously As Boolean, intIndex As Integer
    
    ' See if the form has already called this sub
    ' If not, add to strFormsCentered()
    boolSubCalledPreviously = False
    For intIndex = 0 To intFormsCenteredCount - 1
        If strFormsCentered(intIndex) = frmThisForm.Name Then
            boolSubCalledPreviously = True
            Exit For
        End If
    Next intIndex
    
    If Not boolSubCalledPreviously Then
        ' First time sub called this sub
        ' Add to strFormsCentered()
        If intFormsCenteredCount < MAX_RESIZE_FORMS_TO_REMEMBER Then
            intFormsCenteredCount = intFormsCenteredCount + 1
            strFormsCentered(intFormsCenteredCount - 1) = frmThisForm.Name
        Else
            Debug.Assert False
        End If
    End If
    
    ' If form called previously and boolSizeAndCenterOnlyOncePerProgramSessionis true, then exit sub
    If boolSizeAndCenterOnlyOncePerProgramSession And boolSubCalledPreviously Then
        Exit Sub
    End If
    
    ' Resize Window
    With frmThisForm
        .WindowState = vbNormal
        If lngWindowWidth > 0 Then .Width = lngWindowWidth
        If lngWindowHeight > 0 Then .Height = lngWindowHeight
    End With
    
    ' Assume the first form loaded is the main form
    ' May need to be customized if ported to other applications
    Set frmMainAppForm = Forms(0)
    
    ' Find the desktop area (width and height)
    lngWindowAreaWidth = Screen.Width
    lngWindowAreaHeight = Screen.Height
    
    ' Check the aspect ratio of WindowAreaWidth / WindowAreaHeight
    If lngWindowAreaHeight > 0 Then
        dblAspectRatio = lngWindowAreaWidth / lngWindowAreaHeight
    Else
        dblAspectRatio = 1.333
    End If
    
    ' Typical desktop areas and aspect ratios
    ' Normal Desktops have aspect ratios of 1.33 or 1.5
    ' HDTV desktops have an aspect ratio of 1.6 or 1.7
    ' Horizontal Dual Monitors have an aspect ratio of 2.66 or 2.5
    ' Vertical Dual Monitors have an aspectr ratio of 0.67 or 0.62
    
    ' Determine if using dual monitors
    If dblAspectRatio < 1 Or dblAspectRatio > 2 Then
        boolDualMonitor = True
        If dblAspectRatio > 2 Then
            ' Aspect ratio greater than 2 - using horizontal dual monitors
            boolHorizontalDual = True
            lngWorkingAreaWidth = Screen.Width / 2
            lngWorkingAreaHeight = Screen.Height
            
            If frmMainAppForm.Left > lngWorkingAreaWidth Then
                ' Main app window on second monitor
                ' Set intDualMonitorToUse if not explicitly set
                If intDualMonitorToUse < 0 Then
                    intDualMonitorToUse = 1
                End If
            End If
        Else
            ' Aspect ratio must be less than 1 - using vertical dual monitors
            boolHorizontalDual = False
            lngWorkingAreaWidth = Screen.Width
            lngWorkingAreaHeight = Screen.Height / 2
            
            If frmMainAppForm.Top > lngWorkingAreaHeight Then
                ' Main app window on second monitor
                ' Set intDualMonitorToUse if not explicitly set
                If intDualMonitorToUse < 0 Then
                    intDualMonitorToUse = 1
                End If
            End If
        End If
    Else
        ' Aspect ratio between 1 and 2
        ' Using a single monitor
        boolDualMonitor = False
        lngWorkingAreaWidth = Screen.Width
        lngWorkingAreaHeight = Screen.Height
    End If
    
    With frmThisForm
        ' Position window
        Select Case intCenterMode
        Case cWindowUpperThird
            lngWindowLeftToSet = (lngWorkingAreaWidth - .Width) \ 2
            lngWindowTopToSet = (lngWorkingAreaHeight - .Height) \ 3
        Case cWindowLowerThird
            lngWindowLeftToSet = (lngWorkingAreaWidth - .Width) \ 2
            lngWindowTopToSet = (lngWorkingAreaHeight - .Height) * 2 \ 3
        Case cWindowMiddleLeft
            lngWindowLeftToSet = 0
            lngWindowTopToSet = (lngWorkingAreaHeight - .Height) \ 2
        Case cWindowMiddleRight
            lngWindowLeftToSet = lngWorkingAreaWidth - .Width
            lngWindowTopToSet = (lngWorkingAreaHeight - .Height) \ 2
        Case cWindowTopCenter
            lngWindowLeftToSet = (lngWorkingAreaWidth - .Width) \ 2
            lngWindowTopToSet = 0
        Case cWindowBottomCenter
            lngWindowLeftToSet = (lngWorkingAreaWidth - .Width) \ 2
            lngWindowTopToSet = lngWorkingAreaHeight - .Height - 500
        Case cWindowBottomRight
            lngWindowLeftToSet = lngWorkingAreaWidth - .Width
            lngWindowTopToSet = lngWorkingAreaHeight - .Height - 500
        Case cWindowBottomLeft
            lngWindowLeftToSet = 0
            lngWindowTopToSet = lngWorkingAreaHeight - .Height - 500
        Case cWindowTopRight
            lngWindowLeftToSet = lngWorkingAreaWidth - .Width
            lngWindowTopToSet = 0
        Case cWindowTopLeft
            lngWindowLeftToSet = 0
            lngWindowTopToSet = 0
        Case Else ' Includes cWindowExactCenter = 0
            lngWindowLeftToSet = (lngWorkingAreaWidth - .Width) \ 2
            lngWindowTopToSet = (lngWorkingAreaHeight - .Height) \ 2
        End Select
        
        ' Move to second monitor if explicitly stated or if the main window is already on the second monitor
        If boolDualMonitor And intDualMonitorToUse > 0 Then
            ' Place window on second monitor
            If boolHorizontalDual Then
                ' Horizontal dual - Shift to the right
                lngWindowLeftToSet = lngWindowLeftToSet + lngWorkingAreaWidth
            Else
                ' Vertical dual - Shift down
                lngWindowTopToSet = lngWindowTopToSet + lngWorkingAreaHeight
            End If
        End If
        
        ' Actually position the window
        .Move lngWindowLeftToSet, lngWindowTopToSet
    End With

End Sub

Public Function SpacePad(strWork As String, intLength As Integer) As String
    ' Adds spaces to strWork until the length = intLength
    
    Do While Len(strWork) < intLength
        strWork = strWork & " "
    Loop
    SpacePad = strWork
End Function

Public Function SpacePadToFront(strWork As String, intLength As Integer) As String
    ' Adds spaces to the beginning of strWork until the length = intLength
    
    Do While Len(strWork) < intLength
        strWork = " " & strWork
    Loop
    SpacePadToFront = strWork
End Function

Public Function StringToNumber(ByVal strWork As String, Optional ByRef intNumLength As Integer, Optional ByRef intErrorCode As Integer, Optional blnAllowMinusSign As Boolean = False, Optional blnAllowPlusSign As Boolean = False, Optional blnAllowESymbol As Boolean = False, Optional blnMultipleDecimalPointIsError As Boolean = True, Optional ByVal strDecimalPointSymbol As String = ".") As Double
    ' Looks for a number at the start of strWork and returns it if found
    ' strWork can contain non-numeric characters after the number; only the number will be returned
    ' intNumLength returns the length of the number, including the decimal point and any negative sign or E symbol
    ' When blnAllowESymbol = True, then looks for exponential numbers, like 3.23E+04 or 2.48E-084
    ' If an error is found or no number is present, then 0 is returned and intNumLength is set to 0, and intErrorCode is assigned the error code:
    '   0 = No Error
    '  -1 = No number
    '  -3 = No number at all or (more likely) no number after decimal point
    '  -4 = More than one decimal point
    
    ' Examples:
    '  23Text           Returns 23 and intNumLength = 2
    '  23.432Text       Returns 23.432 and intNumLength = 6
    '  .3Text           Returns 0.3 and intNumLength = 2
    '  0.3Text          Returns 0.3 and intNumLength = 3
    '  3Text            Returns 3 and intNumLength = 1
    '  3.Text           Returns 3 and intNumLength = 2
    '  Text             Returns 0 and intNumLength = 0 and intErrorCode = -3
    '  .Text            Returns 0 and intNumLength = 0 and intErrorCode = -3
    '  4.23.Text        Returns 0 and intNumLength = 0 and intErrorCode = -4  (must have blnMultipleDecimalPointIsError = True)
    '  -43Text          Returns -43 and intNumLength = 2            (must have blnAllowMinusSign = True)
    '  32E+48Text       Returns 32E+48 and intNumLength = 6         (must have blnAllowESymbol = True)
    
    Dim strTestChar As String
    Dim strFoundNum As String, intIndex As Integer, intDecPtCount As Integer
    Dim blnNumberFound As Boolean, blnESymbolFound As Boolean
    
    If strDecimalPointSymbol = "" Then
        strDecimalPointSymbol = DetermineDecimalPoint()
    End If
    
    ' Set intNumLength to -1 for now
    ' If it doesn't get set to 0 (due to an error), it will get set to the
    '   length of the matched number before exiting the sub
    intNumLength = -1
    
    If Len(strWork) > 0 Then
        strFoundNum = Left(strWork, 1)
        If IsNumeric(strFoundNum) Then
            blnNumberFound = True
        ElseIf strFoundNum = strDecimalPointSymbol Then
            blnNumberFound = True
            intDecPtCount = intDecPtCount + 1
        ElseIf (strFoundNum = "-" And blnAllowMinusSign) Then
            blnNumberFound = True
        ElseIf (strFoundNum = "+" And blnAllowPlusSign) Then
            blnNumberFound = True
        End If
    End If
    
    If blnNumberFound Then
        ' Start of string is a number or a decimal point, or (if allowed) a negative or plus sign
        ' Continue looking
        
        intIndex = 2
        Do While intIndex <= Len(strWork)
            strTestChar = Mid(strWork, intIndex, 1)
            If IsNumeric(strTestChar) Then
                strFoundNum = strFoundNum & strTestChar
            ElseIf strTestChar = strDecimalPointSymbol Then
                intDecPtCount = intDecPtCount + 1
                If intDecPtCount = 1 Then
                    strFoundNum = strFoundNum & strTestChar
                Else
                    Exit Do
                End If
            ElseIf (UCase(strTestChar) = "E" And blnAllowESymbol) Then
                ' E symbol found; only add to strFoundNum if followed by a + and a number,
                '                                                        a - and a number, or another number
                strTestChar = Mid(strWork, intIndex + 1, 1)
                If IsNumeric(strTestChar) Then
                    strFoundNum = strFoundNum & "E" & strTestChar
                    intIndex = intIndex + 2
                    blnESymbolFound = True
                ElseIf strTestChar = "+" Or strTestChar = "-" Then
                    If IsNumeric(Mid(strWork, intIndex + 2, 1)) Then
                        strFoundNum = strFoundNum & "E" & strTestChar & Mid(strWork, intIndex + 2, 1)
                        intIndex = intIndex + 3
                        blnESymbolFound = True
                    End If
                End If
                
                If blnESymbolFound Then
                    ' Continue looking for numbers after the E symbol
                    ' However, only allow pure numbers; not + or - or .
                    
                    Do While intIndex <= Len(strWork)
                        strTestChar = Mid(strWork, intIndex, 1)
                        If IsNumeric(strTestChar) Then
                            strFoundNum = strFoundNum & strTestChar
                        ElseIf strTestChar = strDecimalPointSymbol Then
                            If blnMultipleDecimalPointIsError Then
                                ' Set this to 2 to force the multiple decimal point error to appear
                                intDecPtCount = 2
                            End If
                            Exit Do
                        Else
                            Exit Do
                        End If
                        intIndex = intIndex + 1
                    Loop
                End If
                
                Exit Do
            Else
                Exit Do
            End If
            intIndex = intIndex + 1
        Loop
        
        If intDecPtCount > 1 And blnMultipleDecimalPointIsError Then
            ' Too many decimal points
            intNumLength = 0          ' No number found
            intErrorCode = -4
            StringToNumber = 0
        ElseIf Len(strFoundNum) = 0 Or strFoundNum = strDecimalPointSymbol Then
            ' No number at all or (more likely) no number after decimal point
            intNumLength = 0          ' No number found
            intErrorCode = -3
            StringToNumber = 0
        Else
            ' All is fine
            intNumLength = Len(strFoundNum)
            intErrorCode = 0
            StringToNumber = CDblSafe(strFoundNum)
        End If
    Else
        intNumLength = 0          ' No number found
        intErrorCode = -1
        StringToNumber = 0
    End If
    
End Function

Public Function StripChrZero(ByRef DataString As String) As String
    Dim intCharLoc As Integer
    
    intCharLoc = InStr(DataString, Chr(0))
    If intCharLoc > 0 Then
        DataString = Left(DataString, intCharLoc - 1)
    End If
    StripChrZero = DataString
    
End Function

Public Function StripFullPath(ByVal strFilePathIn As String, Optional ByRef strStrippedPath As String) As String
    ' Removes all path info from strFilePathIn, returning just the filename
    ' The path of the file is returned in strStrippedPath (including the trailing \)
    
    Dim fso As New FileSystemObject

    strStrippedPath = fso.GetParentFolderName(strFilePathIn)
    If Right(strStrippedPath, 1) <> "\" And Right(strStrippedPath, 1) <> "/" Then strStrippedPath = strStrippedPath & "\"
    
    StripFullPath = fso.GetFileName(strFilePathIn)
    
    Set fso = Nothing
    
''    Dim intCharLoc As Integer
''
''    intCharLoc = InStrRev(strFilePathIn, "\")
''    If intCharLoc > 0 Then
''        strStrippedPath = Left(strFilePathIn, intCharLoc)
''        StripFullPath = Mid(strFilePathIn, intCharLoc + 1)
''    Else
''        StripFullPath = strFilePathIn
''    End If
    
End Function

Private Sub SwapSingle(ByRef FirstValue As Single, ByRef SecondValue As Single)
    Dim sngTemp As Single
    sngTemp = FirstValue
    FirstValue = SecondValue
    SecondValue = sngTemp
End Sub

Public Sub SwapValues(ByRef FirstValue As Variant, ByRef SecondValue As Variant)
    Dim varTemp As Variant
    varTemp = FirstValue
    FirstValue = SecondValue
    SecondValue = varTemp
End Sub

Public Sub TextBoxKeyPressHandler(txtThisTextBox As TextBox, ByRef KeyAscii As Integer, Optional AllowNumbers As Boolean = True, Optional AllowDecimalPoint As Boolean = False, Optional AllowNegativeSign As Boolean = False, Optional AllowCharacters As Boolean = False, Optional AllowPlusSign As Boolean = False, Optional AllowUnderscore As Boolean = False, Optional AllowDollarSign As Boolean = False, Optional AllowEmailChars As Boolean = False, Optional AllowSpaces As Boolean = False, Optional AllowECharacter As Boolean = False, Optional boolAllowCutCopyPaste As Boolean = True)
    ' Note that the AllowECharacter option has been added to allow the
    '  user to type numbers in scientific notation
    
    ' Checks KeyAscii to see if it's valid
    ' If it isn't, it is set to 0
    
    Select Case KeyAscii
    Case 1
        ' Ctrl+A -- Highlight entire text box
        txtThisTextBox.SelStart = 0
        txtThisTextBox.SelLength = Len(txtThisTextBox.Text)
        KeyAscii = 0
    Case 24, 3, 22
        ' Cut, copy, paste, or delete was pressed; let the command occur
        If Not boolAllowCutCopyPaste Then KeyAscii = 0
    Case 26
        ' Ctrl+Z = Undo
        KeyAscii = 0
        txtThisTextBox.Text = mTextBoxValueSaved
    Case 8
        ' Backspace is allowed
    Case 48 To 57: If Not AllowNumbers Then KeyAscii = 0
    Case 32: If Not AllowSpaces Then KeyAscii = 0
    Case 36: If Not AllowDollarSign Then KeyAscii = 0
    Case 43: If Not AllowPlusSign Then KeyAscii = 0
    Case 45: If Not AllowNegativeSign Then KeyAscii = 0
    Case 44, 46:
        Select Case glbDecimalSeparator
        Case ","
            If KeyAscii = 46 Then KeyAscii = 0
        Case Else   ' includes "."
            If KeyAscii = 44 Then KeyAscii = 0
        End Select
        If Not AllowDecimalPoint Then KeyAscii = 0
    Case 64: If Not AllowEmailChars Then KeyAscii = 0
    Case 65 To 90, 97 To 122
        If Not AllowCharacters Then
            If Not AllowECharacter Then
                KeyAscii = 0
            Else
                If KeyAscii = 69 Or KeyAscii = 101 Then
                    KeyAscii = 69
                Else
                    KeyAscii = 0
                End If
            End If
        End If
    Case 95: If Not AllowUnderscore Then KeyAscii = 0
    Case Else
        KeyAscii = 0
    End Select

End Sub

Public Sub TextBoxGotFocusHandler(txtThisTextBox As VB.TextBox, Optional blnSelectAll As Boolean = True)
    ' Selects the text in the given textbox if blnSelectAll = true
    ' Stores the current textbox value in mTextBoxValueSaved
    
    If blnSelectAll Then
        txtThisTextBox.SelStart = 0
        txtThisTextBox.SelLength = Len(txtThisTextBox.Text)
    End If
    
    SetMostRecentTextBoxValue txtThisTextBox.Text
End Sub

Public Function TrimFileName(strFilePath As String) As String
    ' Examines strFilePath, looking from the right for a \
    ' If found, returns only the portion after \
    ' Otherwise, returns the enter string
    Dim lngLastSlashLoc As Long
    Dim strTrimmedPath As String
    
    lngLastSlashLoc = InStrRev(strFilePath, "\")
    
    If lngLastSlashLoc > 0 Then
        strTrimmedPath = Mid(strFilePath, lngLastSlashLoc + 1)
    Else
        strTrimmedPath = strFilePath
    End If
    
    TrimFileName = strTrimmedPath
End Function

Public Sub UnloadAllForms(strCallingFormName As String)

    Dim frmThisForm As VB.Form
    ' Unload all the other forms
    For Each frmThisForm In Forms
        If frmThisForm.Name <> strCallingFormName Then
            Unload frmThisForm
        End If
    Next

    ' If all the other forms did not unload, then the following will be false and
    ' the program will break.  Force the application to end.
    Debug.Assert Forms.Count = 1
    If Forms.Count > 1 Then
        'End
    End If

End Sub

Public Function ValidateDualTextBoxes(txtFirstTextBox As TextBox, txtSecondTextBox As TextBox, boolFavorFirstTextBox As Boolean, dblLowerBound As Double, dblUpperBound As Double, Optional dblDefaultSeparationAmount As Double = 1) As Boolean
    
    ' Makes sure txtFirstTextBox is less than or equal to txtSecondTextBox
    ' Returns True if all is OK; returns False if one of the textboxes had to be corrected
    
    Dim blnTextboxUpdated As Boolean
    
    If dblUpperBound > 0 Then
        If boolFavorFirstTextBox Then
            If Val(txtFirstTextBox) > Val(txtSecondTextBox) Then
                txtSecondTextBox = CStr(Val(txtFirstTextBox) + dblDefaultSeparationAmount)
                blnTextboxUpdated = True
            End If
        Else
            If Val(txtFirstTextBox) > Val(txtSecondTextBox) Then
                txtFirstTextBox = CStr(Val(txtSecondTextBox) - dblDefaultSeparationAmount)
                blnTextboxUpdated = True
            End If
        End If
        
        If Val(txtSecondTextBox) > dblUpperBound Then
            txtSecondTextBox = Format(dblUpperBound, "0.00")
            blnTextboxUpdated = True
        End If
        If Val(txtFirstTextBox) < dblLowerBound Then
            txtFirstTextBox = Format(dblLowerBound, "0.00")
            blnTextboxUpdated = True
        End If

        If boolFavorFirstTextBox Then
            blnTextboxUpdated = Not ValidateDualTextBoxes(txtFirstTextBox, txtSecondTextBox, False, dblLowerBound, dblUpperBound)
        End If
    End If

    ValidateDualTextBoxes = Not blnTextboxUpdated
    
End Function

Public Function ValidateTextboxValueLng(txtThisTextControl As TextBox, lngMinimumVal As Long, lngMaximumVal As Long, lngdefaultVal As Long) As Long
    If Val(txtThisTextControl) < lngMinimumVal Or Val(txtThisTextControl) > lngMaximumVal Or Not IsNumeric(txtThisTextControl) Then
        txtThisTextControl = Trim(Str(lngdefaultVal))
    End If
    ValidateTextboxValueLng = Val(txtThisTextControl)
End Function

Public Function ValidateTextboxValueDbl(txtThisTextControl As TextBox, dblMinimumVal As Double, dblMaximumVal As Double, dbldefaultVal As Double) As Double
    If Val(txtThisTextControl) < dblMinimumVal Or Val(txtThisTextControl) > dblMaximumVal Or Not IsNumeric(txtThisTextControl) Then
        txtThisTextControl = Trim(Str(dbldefaultVal))
    End If
    ValidateTextboxValueDbl = Val(txtThisTextControl)
End Function

Public Function ValidateValueDbl(ByRef dblThisValue As Double, dblMinimumVal As Double, dblMaximumVal As Double, dbldefaultVal As Double) As Double
    If dblThisValue < dblMinimumVal Or dblThisValue > dblMaximumVal Then
        dblThisValue = dbldefaultVal
    End If
    
    ValidateValueDbl = dblThisValue
End Function

Public Function ValidateValueLng(ByRef lngThisValue As Long, lngMinimumVal As Long, lngMaximumVal As Long, lngdefaultVal As Long) As Long
    If lngThisValue < lngMinimumVal Or lngThisValue > lngMaximumVal Then
        lngThisValue = lngdefaultVal
    End If
    
    ValidateValueLng = lngThisValue
End Function

Private Function VerifySort(sngArray() As Single, lngLowIndex As Long, lngHighIndex As Long) As Boolean
    Dim blnInOrder As Boolean
    Dim lngIndex As Long
    
    blnInOrder = True
    
    For lngIndex = lngLowIndex To lngHighIndex - 1
        If sngArray(lngIndex) > sngArray(lngIndex + 1) Then
            blnInOrder = False
            Exit For
        End If
    Next lngIndex
    
    VerifySort = blnInOrder
End Function

Public Sub VerifyValidWindowPos(frmThisForm As VB.Form, Optional lngMinWidth As Long = 500, Optional lngMinHeight As Long = 500, Optional MinVisibleFormArea As Long = 500)
    ' Make sure the window isn't too small and is visible on the desktop
    
    Dim lngReturn As Long
    Dim lngScreenWidth As Long, lngScreenHeight As Long
    
    lngReturn = GetDesktopSize(lngScreenHeight, lngScreenWidth, True)
    
    If lngScreenHeight < Screen.Height Then lngScreenHeight = Screen.Height
    If lngScreenWidth < Screen.Width Then lngScreenWidth = Screen.Width
    
    On Error GoTo VerifyValidWindowPosErrorHandler
    With frmThisForm
        If .WindowState = vbMinimized Then
            .WindowState = vbNormal
        End If
        
        If .Width < lngMinWidth Then .Width = lngMinWidth
        If .Height < lngMinHeight Then .Height = lngMinHeight
                
        If .Left > lngScreenWidth - MinVisibleFormArea Or _
           .Top > lngScreenHeight - MinVisibleFormArea Or _
           .Left < 0 Or .Top < 0 Then
           SizeAndCenterWindow frmThisForm, cWindowUpperThird, .Width, .Height, False
        End If
    End With
    
    Exit Sub
    
VerifyValidWindowPosErrorHandler:
    ' An error occured
    ' The form is probably minimized; we'll ignore it
    Debug.Print "Error occured in VerifyValidWindowPos: " & Err.Description
    
End Sub

Public Sub WindowStayOnTop(hwnd As Long, boolStayOnTop As Boolean, Optional lngFormPosLeft As Long = 0, Optional lngFormPosTop As Long = 0, Optional lngFormPosWidth As Long = 600, Optional lngFormPosHeight As Long = 500)
    ' Toggles the behavior of the given window to "stay on top" of all other windows
    ' The new form sizes (lngFormPosLeft, lngFormPosTop, lngFormPosWidth, lngFormPosHeight)
    '  are in pixels
    
    Dim lngTopMostSwitch As Long
    
    If boolStayOnTop Then
        ' Turn on the TopMost attribute.
        lngTopMostSwitch = conHwndTopmost
    Else
        ' Turn off the TopMost attribute.
        lngTopMostSwitch = conHwndNoTopmost
    End If
    
    SetWindowPos hwnd, lngTopMostSwitch, lngFormPosLeft, lngFormPosTop, lngFormPosWidth, lngFormPosHeight, conSwpNoActivate Or conSwpShowWindow
End Sub

Public Function YesNoBox(strMessage As String, strTitle As String) As VbMsgBoxResult
    ' Displays a Message Box with OK/Cancel buttons (i.e. yes/no)
    ' uses vbDefaultButton2 to make sure No or Cancel is the default button
    
    Dim DialogType As Integer

    ' The dialog box should have Yes and No buttons,
    ' and a question icon.
    DialogType = vbYesNo + vbQuestion + vbDefaultButton2

    ' Display the dialog box and get user's response.
    YesNoBox = MsgBox(strMessage, DialogType, strTitle)
End Function

