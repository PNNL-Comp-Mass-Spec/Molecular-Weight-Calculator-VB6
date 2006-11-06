VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.1#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H80000005&
   Caption         =   "RTFbox"
   ClientHeight    =   6576
   ClientLeft      =   2316
   ClientTop       =   1800
   ClientWidth     =   6696
   LinkTopic       =   "Form1"
   ScaleHeight     =   6576
   ScaleWidth      =   6696
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   6100
      Left            =   336
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   240
      Width           =   6132
      _ExtentX        =   10816
      _ExtentY        =   10753
      _Version        =   327681
      BackColor       =   16777215
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      Appearance      =   0
      RightMargin     =   1
      OLEDragMode     =   0
      OLEDropMode     =   0
      FileName        =   "C:\sample.rtf"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuPop1 
      Caption         =   ""
      NegotiatePosition=   1  'Left
      Visible         =   0   'False
      Begin VB.Menu define 
         Caption         =   "Definition"
      End
      Begin VB.Menu note 
         Caption         =   "Note"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Declare Function HideCaret Lib "user32" (ByVal hwnd As Long) As Long

Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202

Dim xy As Long
Dim Char As String
Dim Cnt As Integer
Dim Pos As Long
Dim ItsALetter As Boolean
Dim Word As String

Private Sub define_Click()
    Word = Left$(Word, Cnt - 1)
    MsgBox (Word)
End Sub

Private Sub Form_Load()
    If Dir("C:\sample.rtf") = "" Then MsgBox ("The program looks for the SAMPLE.RTF file in the root C:\ directory."): End
End Sub

Private Sub note_Click()
    
    'Note: This isn't how I'm going to implement the hypertext, but it just shows that it works.
    'Also note that the formatting doesn't have any effect on the SelStart property. It's the
    'actual number of characters from the start of the box, starting with "0", though, instead of "1"
    'i.e., the first word's SelStart is "0"
    
    Select Case RichTextBox1.SelStart
        
        Case 144
        msg = "You clicked on --distribution-- in the third paragraph"
       
        Case 775
        msg = "You clicked on the second instance of --commercial-- in the fourth paragraph."
        
        Case Else
        msg = "SelStart for the word you clicked is" & Str(RichTextBox1.SelStart)
    
    End Select
         
    MsgBox (msg)
    
End Sub

Private Sub RichTextBox1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

'Remove highlight if present
        If Cnt > 0 Then
            RichTextBox1.SelLength = Cnt - 1
            RichTextBox1.SelColor = RGB(0, 0, 0)
            RichTextBox1.SelLength = 0
        End If
End Sub

Private Sub RichTextBox1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Button = 2 Then     'right mouse button
        
        'Delay window repainting. Trying to use this to eliminate window flashing... not totally working.
        Call LockWindowUpdate(RichTextBox1.hwnd)
        
        '(Re-)initialize some variables
        Word = ""
        Cnt = 0
        ItsALetter = True
                
        'move insertion point
        b = CLng(x / Screen.TwipsPerPixelX)
        c = CLng(y / Screen.TwipsPerPixelY)
        xy = CLng((b) + ((c) * (2 ^ 16)))
        a = SendMessage(RichTextBox1.hwnd, WM_LBUTTONDOWN, 0, ByVal xy)
        a = SendMessage(RichTextBox1.hwnd, WM_LBUTTONUP, 0, ByVal xy)
        
        'go to the beginning of the selected word
        RichTextBox1.UpTo "AaBbCcDdEeFfGgHhIiJjKkLlMmNnOoPpQqRrSsTtUuVvWwXxYyZz", False, True
        
        'See if we're in the excluded zone
        If (RichTextBox1.SelStart < 106) Then Exit Sub
        
        'Count up all the letters in the word
        Pos = RichTextBox1.SelStart + 1
        Do While ItsALetter
            Char = Mid$(RichTextBox1.Text, Pos, 1)
            ItsALetter = ((Asc(Char) >= 65 And Asc(Char) <= 90) Or (Asc(Char) >= 97 And Asc(Char) <= 122))
            
            'Note: this isn't necessary in this routine, but it's nice to have if needed elsewhere
            'You just need to strip the last character using  Left$(Word, Cnt-1) when the routine
            'finishes to have the selected word, as we do in the "Definition" submenu
            Word = Word & Char
            
            'Check next letter
            Cnt = Cnt + 1
            Pos = Pos + 1
        Loop
        
        'Highlight word using color red
        RichTextBox1.SelLength = Cnt - 1
        RichTextBox1.SelColor = RGB(255, 0, 0)
        RichTextBox1.SelLength = 0
        
        'Hide the caret
        Call HideCaret(RichTextBox1.hwnd)
        
        'Repaint Screen
        Call LockWindowUpdate(0)
                        
        'Pop up the menu
        PopupMenu mnuPop1, vbPopupMenuCenterAlign, RichTextBox1.Left + x - 100, RichTextBox1.Top + y + 160
        
        'Hide the caret again
        Call HideCaret(RichTextBox1.hwnd)
        
    End If
        
End Sub

