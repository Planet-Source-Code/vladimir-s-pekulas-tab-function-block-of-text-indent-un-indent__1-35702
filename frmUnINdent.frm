VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmUnINdent 
   Caption         =   "Form1"
   ClientHeight    =   5175
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7710
   LinkTopic       =   "Form1"
   ScaleHeight     =   5175
   ScaleWidth      =   7710
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox RTF 
      Height          =   4335
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   7646
      _Version        =   393217
      Enabled         =   -1  'True
      FileName        =   "C:\Documents and Settings\Administrator\Desktop\1\TabFunction\sample_file.txt"
      TextRTF         =   $"frmUnINdent.frx":0000
   End
End
Attribute VB_Name = "frmUnINdent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
  

'// CAPTURE KEY UP EVENT
Private Sub RTF_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 9 And Shift = 1 Then
        If RTF.SelLength > 0 Then
            Call UN_INDENT_TEXT '// un-indent block of selected text
            KeyCode = 27        '// Cancel the incoming tab key
        End If
    End If
End Sub


'// UN-INDENT BLOCK OF SELECTED TEXT
Function UN_INDENT_TEXT()
Dim intSTART As Integer, strOriginal As String
    intSTART = RTF.SelStart
    strOriginal = RTF.SelText
    
    ONE_LINE = split(strOriginal, vbCrLf)
    
    For I = 1 To UBound(ONE_LINE)
            
        
    Next I
    

End Function




