VERSION 5.00
Begin VB.MDIForm MDI 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000C&
   Caption         =   "MDI Parent Window"
   ClientHeight    =   4050
   ClientLeft      =   4725
   ClientTop       =   2055
   ClientWidth     =   5670
   LinkTopic       =   "MDIForm1"
   Begin VB.PictureBox picTaskbar 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   0
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   378
      TabIndex        =   0
      Top             =   3705
      Width           =   5670
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFile_New 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFile_Close 
         Caption         =   "&Close"
         Shortcut        =   ^{F4}
      End
      Begin VB.Menu mnuFile_Sep01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_Exit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "MDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
    InitTaskbar
End Sub

Private Sub MDIForm_Resize()
    DrawTaskbar picTaskbar
End Sub


Private Sub mnuFile_Exit_Click()
    Unload Me
    End
End Sub


Private Sub mnuFile_New_Click()
    Dim NewForm As New frmChild
    NewForm.Show
    NewForm.Caption = "Child Window " & windowCount + 1
    SetWindowCount
    DrawTaskbar picTaskbar
End Sub


Private Sub picTaskbar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim which As Long
    which = GetItemFromXY(CLng(X), CLng(Y))
    If which = downWhich Then Exit Sub
    
    downWhich = which
    DrawTaskbar picTaskbar
End Sub

Private Sub picTaskbar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim which As Long
    which = GetItemFromXY(CLng(X), CLng(Y))
    If which = overWhich Then Exit Sub
    
    overWhich = which
    DrawTaskbar picTaskbar
End Sub

Private Sub picTaskbar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetFocusByIndex downWhich
    
    If downWhich <> -1 Then activeWhich = downWhich
    downWhich = -1
    DrawTaskbar picTaskbar
End Sub

Private Sub picTaskbar_Paint()
    DrawTaskbar picTaskbar
End Sub


