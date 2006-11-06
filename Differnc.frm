VERSION 5.00
Begin VB.Form frmDiff 
   Caption         =   "Percent Solver Differences"
   ClientHeight    =   4470
   ClientLeft      =   720
   ClientTop       =   2355
   ClientWidth     =   6645
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
   HelpContextID   =   3030
   Icon            =   "Differnc.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4470
   ScaleWidth      =   6645
   Tag             =   "8600"
   Begin VB.CommandButton cmdCopy 
      Caption         =   "&Copy"
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
      Left            =   1200
      TabIndex        =   1
      Tag             =   "8610"
      ToolTipText     =   "Copies the results to the clipboard"
      Top             =   3960
      Width           =   1035
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
      Left            =   2520
      TabIndex        =   0
      Tag             =   "4000"
      Top             =   3960
      Width           =   1035
   End
   Begin VB.Label lblDiff3 
      Height          =   3615
      Left            =   4440
      TabIndex        =   4
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblDiff2 
      Height          =   3615
      Left            =   2280
      TabIndex        =   3
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label lblDiff1 
      Height          =   3615
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmDiff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCopy_Click()
    ReDim Diff(3) As String
    Dim strTextToCopy As String
    Dim intCharLoc As Integer, intIndex As Integer
    ReDim start(3) As Integer
    
    ' Copy results to Clipboard
    Clipboard.Clear
    
    Diff(1) = lblDiff1.Caption
    Diff(2) = lblDiff2.Caption
    Diff(3) = lblDiff3.Caption
    start(1) = 1
    start(2) = 1
    start(3) = 1
    strTextToCopy = frmMain.rtfFormulaSingle.Text & vbCrLf & frmMain.lblValueForX.Caption
    
    ' Format results into columns
    Do
        For intIndex = 1 To 3
            For intCharLoc = start(intIndex) To Len(Diff(intIndex))
                If Asc(Mid(Diff(intIndex), intCharLoc, 1)) = 13 Or Asc(Mid(Diff(intIndex), intCharLoc, 1)) = 10 Then
                    strTextToCopy = strTextToCopy & Trim(Mid(Diff(intIndex), start(intIndex), intCharLoc - start(intIndex) - 0))
                    start(intIndex) = intCharLoc + 1
                    Exit For
                End If
                If intCharLoc = Len(Diff(intIndex)) Then
                    strTextToCopy = strTextToCopy & Trim(Mid(Diff(intIndex), start(intIndex), intCharLoc - start(intIndex) + 1))
                    start(intIndex) = Len(Diff(intIndex))
                    Exit For
                End If
            Next intCharLoc
            If intIndex = 1 Or intIndex = 2 Then strTextToCopy = strTextToCopy & vbTab
        Next intIndex
        strTextToCopy = strTextToCopy & vbCrLf
    Loop Until start(1) >= Len(Diff(1))

    Clipboard.SetText strTextToCopy, vbCFText

End Sub

Private Sub cmdOK_Click()
    frmDiff.Hide
End Sub

Private Sub cmdOK_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyEscape Then cmdOK_Click
End Sub

Private Sub Form_Activate()
    ' Put window in center of screen
    SizeAndCenterWindow Me, cWindowLowerThird, 8000, 5000

End Sub

Private Sub Form_Load()
                    
    Form_Resize
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    QueryUnloadFormHandler Me, Cancel, UnloadMode
End Sub

Private Sub Form_Resize()
    
    If frmDiff.WindowState <> vbMinimized Then
        If frmDiff.Width < 4500 Then frmDiff.Width = 4500
        If frmDiff.Height < 2500 Then frmDiff.Height = 2500
        
        With lblDiff1
            .top = 240
            .Left = 240
            .Height = frmDiff.ScaleHeight * 0.9
            .Width = (frmDiff.ScaleWidth - 50) / 4
        End With
        
        With lblDiff2
            .top = 240
            .Left = frmDiff.ScaleWidth / 3
            .Height = frmDiff.ScaleHeight * 0.9
            .Width = (frmDiff.ScaleWidth - 50) / 4
        End With
        
        With lblDiff3
            .top = 240
            .Left = 2 * frmDiff.ScaleWidth / 3
            .Height = frmDiff.ScaleHeight * 0.9
            .Width = (frmDiff.ScaleWidth - 50) / 4
        End With
        
        cmdCopy.top = frmDiff.ScaleHeight - cmdOK.Height - 100
        cmdCopy.Left = (frmDiff.Width - cmdOK.Width) / 3
        cmdOK.top = frmDiff.ScaleHeight - cmdOK.Height - 100
        cmdOK.Left = 2 * (frmDiff.Width - cmdOK.Width) / 3
        
    End If
End Sub

