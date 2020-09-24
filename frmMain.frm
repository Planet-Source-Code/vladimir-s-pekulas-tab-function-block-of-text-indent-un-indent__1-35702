VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tab Spacing Sample"
   ClientHeight    =   4215
   ClientLeft      =   150
   ClientTop       =   390
   ClientWidth     =   7230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   7560
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame frmFrame 
      Caption         =   "Tab Spacing Function Sample"
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6975
      Begin RichTextLib.RichTextBox RTF 
         Height          =   3495
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   6165
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   3
         RightMargin     =   50000
         TextRTF         =   $"frmMain.frx":0000
      End
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpenFile 
         Caption         =   "&Open File ..."
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuAction 
      Caption         =   "&Action"
      Visible         =   0   'False
      Begin VB.Menu mnuMoveText 
         Caption         =   "&Indent Selected text"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/////////////////////////////////////////////////////////////
'
'    COPYRIGHT (C) 2002 VLADIMIR S. PEKULAS (Europeum.net)
'
'    PURPOSE: To move (indent) block of selected text as one,
'             instead of one line at the time by using a TAB key
'
'    BUGS:
'             So far none have been discovered. (June 10. 2002)
'
'    SUBS/FUNCTION TO LOOK FOR:
'             TAB_MOVE ; RTF_KeyPress ; SWAP_SPACES ; RTF_KeyDown
'             (The rest is just for fun)
'
'/////////////////////////////////////////////////////////////


'// LOAD SAMPLE FILE
Private Sub Form_Load()
    RTF.LoadFile App.Path & "\sample_file.txt"
End Sub
'// UNLOAD THE FORM - MENU COMMAND
Private Sub mnuExit_Click()
    Unload Me
End Sub
'// OPEN FILE - MENU COMMAND
Private Sub mnuOpenFile_Click()
    With CommonDialog
      .DialogTitle = "Open File ..."
      .ShowOpen
      .DefaultExt = "*.txt"
      RTF.LoadFile .FileName
    End With
End Sub



'// CAPTURE WHAT KEY COMBINATION WAS PRESSED
Private Sub RTF_KeyDown(KeyCode As Integer, Shift As Integer)
  If RTF.SelLength > 0 Then            '// IF NOTHING IS SELECTED THEN DO NOTHING
    If KeyCode = 9 And Shift = 1 Then  '// DON'T TUCH OTHER KEYS
       Call Swap_Spaces                '// UNINDENT THE BLOCK OF TEXT
       KeyCode = 27                    '// CANCEL THE TAB KEY
    End If
  End If
End Sub

'// DETERMINE WHAT KEY WAS PRESSED
Private Sub RTF_KeyPress(KeyAscii As Integer)
    If KeyAscii = 9 Then             '// DON'T TUCH OTHER KEYS
       If Not RTF.SelLength = 0 Then '// IF NOTHING IS SELECTED THEN ALLOW FOR NORMAL TAB KEY
            Call TAB_MOVE(True)      '// MOVE TEXT IF TAB IS PRESSED
            KeyAscii = 27            '// CANCEL THE TAB KEY
       End If
    End If
End Sub


'// UN-INDENT BLOC OF TEXT
'// THIS FUNCTION MAKES THE TEXT MOVES AS A BLOCK TO THE LEFT
'// IT CHECK IF WE LINE HAS AT LEAST 5 SPACES TO MOVE
Function Swap_Spaces()
Dim strORIGINAL As String, I As Integer, strNEW As String, intSTART As Integer
    intSTART = RTF.SelStart
    strORIGINAL = RTF.SelText
    arrLINE = Split(strORIGINAL, vbCrLf)
    '// TAKE EACH LINE AND CHECK IF WE HAVE ROOM TO MOVE
    For I = 0 To UBound(arrLINE)
        If Mid(arrLINE(I), 1, 5) = Space(5) Then
            LN = Mid(arrLINE(I), 5, Len(arrLINE(I)))
        Else
            LN = arrLINE(I)
        End If
        '// CREATE NEW BLOCK
        If I = UBound(arrLINE) Then
          strNEW = strNEW & LN
        Else
          strNEW = strNEW & LN & vbCrLf
        End If
    Next I
    '// SET NEW BLOCK OF TEXT
    RTF.SelText = strNEW
    RTF.SelStart = intSTART
    RTF.SelLength = Len(strNEW)
End Function



'// INDENT BLOC OF TEXT
'// THIS FUNCTION MAKES THE TEXT MOVES AS A BLOCK
Function TAB_MOVE(ByRef USE_TAB As Boolean)
Dim intSTART As Integer, intLENGTH As Integer, txtNEW As String
    intSTART = RTF.SelStart     '// REMEMBER THE START POS OF CARRET
    txtORIGINAL = RTF.SelText
    txtNEW = ""
    SPL_TXT = Split(RTF.SelText, vbCrLf)
        For I = 0 To UBound(SPL_TXT)    '// ADD SPACE BEFORE EACH LINE
            If Not I = UBound(SPL_TXT) Then
                txtNEW = txtNEW & Space(5) & SPL_TXT(I) & vbCrLf
            Else
                txtNEW = txtNEW & Space(5) & SPL_TXT(I) '// DON'T ADD LINE BREAK ON LAST LINE
            End If
        Next I
    RTF.SelText = txtNEW
    RTF.SelStart = intSTART
    RTF.SelLength = Len(txtNEW)
End Function









'// SAMPLE INTEGRATION INTO POP-UP MENU
Private Sub RTF_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        If RTF.SelLength = 0 Then
            mnuMoveText.Enabled = False
        Else
            mnuMoveText.Enabled = True
        End If
        Me.PopupMenu mnuAction
    End If
End Sub
Private Sub mnuMoveText_Click()
    Call RTF_KeyPress(9)
End Sub
'// END OF SAMPLE INTEGRATION
