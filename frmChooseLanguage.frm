VERSION 5.00
Begin VB.Form frmChooseLanguage 
   Caption         =   "Choose Language"
   ClientHeight    =   3345
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4725
   ControlBox      =   0   'False
   HelpContextID   =   4003
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   4725
   StartUpPosition =   3  'Windows Default
   Tag             =   "8400"
   Begin VB.ListBox lstAvailableLanguages 
      Height          =   1620
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   4215
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   360
      Left            =   960
      TabIndex        =   3
      Tag             =   "4010"
      Top             =   2880
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   360
      Left            =   2280
      TabIndex        =   4
      Tag             =   "4020"
      Top             =   2880
      Width           =   1035
   End
   Begin VB.Label lblNoLanguageFiles 
      Caption         =   "No language files are available.  Visit the author's homepage to download alternate languages."
      Height          =   735
      Left            =   360
      TabIndex        =   1
      Tag             =   "8420"
      Top             =   240
      Width           =   4215
   End
   Begin VB.Label lblChooseLanguage 
      Caption         =   "Available languages are shown below.  Please choose the language you wish to use."
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Tag             =   "8410"
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmChooseLanguage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Form-wide global variables
Private intLanguageFilesFound As Integer
Private strLanguageFileInfo(MAX_LANGUAGE_FILE_COUNT, 2) As String

Private Sub PositionFormControls()
    Me.Caption = LookupLanguageCaption(8400, "Choose Language")
    cmdOK.Caption = LookupLanguageCaption(4010, "&Ok")
    cmdCancel.Caption = LookupLanguageCaption(4020, "&Cancel")
    
    lblChooseLanguage.Top = 120
    lblChooseLanguage.Left = 120
    lblChooseLanguage.Caption = LookupLanguageCaption(8410, "Available languages are shown below.  Please choose the language you wish to use.")
    
    lblNoLanguageFiles.Top = lblChooseLanguage.Top
    lblNoLanguageFiles.Left = lblChooseLanguage.Left
    lblNoLanguageFiles.Caption = LookupLanguageCaption(8420, "No language files are available.  Visit the author's homepage to download alternate languages.")
    
    lstAvailableLanguages.Top = 1080
    lstAvailableLanguages.Left = 120

End Sub

Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    Dim boolSuccess As Boolean
    
    If lblNoLanguageFiles.Visible = False Then
        ' If chosen language is different than current language, then load new language
        If lstAvailableLanguages.List(lstAvailableLanguages.ListIndex) <> gCurrentLanguage Then
            ' Change mouse pointer to hourglass
            MousePointer = vbHourglass
            
            boolSuccess = LoadLanguageSettings(strLanguageFileInfo(lstAvailableLanguages.ListIndex, 0), strLanguageFileInfo(lstAvailableLanguages.ListIndex, 1))
            If boolSuccess Then
                SaveSingleDefaultOption "Language", gCurrentLanguage
                SaveSingleDefaultOption "LanguageFile", gCurrentLanguageFileName
                frmMain.lblStatus.ForeColor = vbWindowText
                frmMain.lblStatus.Caption = LookupLanguageCaption(3850, "New default language saved.")
            End If
        
            ' Change mouse pointer back to normal
            MousePointer = vbDefault
        
        End If
    End If
    Me.Hide
    
End Sub

Private Sub Form_Activate()
    ' ReQuery gCurrentPath for Lang*.Ini
    ' Open each one and look for Language=
    ' If Language= exists, display Language value in lstLanguages
    ' Compare to gCurrentLanguage and highlight line with current language if set
    
    ' If no language files are found, hide lblChooseLanguage and show lblNoLanguageFiles
    
    ' User can choose language, and press OK or Cancel
    
    Dim strLanguageFileSearchPath As String, strInputFilePath As String
    Dim strLanguageFileMatch As String, strLineIn As String
    Dim strLanguageName As String, intIndex As Integer
    Dim InFileNum As Integer
    
    ' Position Window
    SizeAndCenterWindow Me, cWindowUpperThird, 4850, 3850
    
    lblChooseLanguage.Visible = True
    lblNoLanguageFiles.Visible = False
    
    strLanguageFileSearchPath = BuildPath(gCurrentPath, "Lang*.Ini")
    
    intLanguageFilesFound = 0
    strLanguageFileMatch = Dir(strLanguageFileSearchPath)
    Do While Len(strLanguageFileMatch) > 0
        strInputFilePath = BuildPath(gCurrentPath, strLanguageFileMatch)
        
        InFileNum = FreeFile()
        Open strInputFilePath For Input As #InFileNum
            Do While Not EOF(InFileNum)
                Line Input #InFileNum, strLineIn
                If LCase(Left(strLineIn, 9)) = "language=" Then
                    strLanguageName = Mid(strLineIn, 10)
                    If Len(strLanguageName) > 0 Then
                        intLanguageFilesFound = intLanguageFilesFound + 1
                        strLanguageFileInfo(intLanguageFilesFound - 1, 0) = strLanguageFileMatch
                        strLanguageFileInfo(intLanguageFilesFound - 1, 1) = strLanguageName
                        Exit Do
                    End If
                End If
            Loop
        Close #InFileNum
        If intLanguageFilesFound >= MAX_LANGUAGE_FILE_COUNT Then Exit Do
        strLanguageFileMatch = Dir
    Loop
    
    If intLanguageFilesFound = 0 Then
        lblChooseLanguage.Visible = False
        lblNoLanguageFiles.Visible = True
    
    Else
        ' Populate List Box
        lstAvailableLanguages.Clear
        For intIndex = 0 To intLanguageFilesFound - 1
            lstAvailableLanguages.AddItem strLanguageFileInfo(intIndex, 1)
            If LCase(gCurrentLanguage) = LCase(strLanguageFileInfo(intIndex, 1)) Then
                lstAvailableLanguages.ListIndex = intIndex
            End If
        Next intIndex
    End If
    
End Sub

Private Sub Form_Load()
    
    PositionFormControls
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    QueryUnloadFormHandler Me, Cancel, UnloadMode
End Sub

Private Sub lstAvailableLanguages_DblClick()
    cmdOK_Click
End Sub
