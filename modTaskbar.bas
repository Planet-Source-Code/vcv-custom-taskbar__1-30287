Attribute VB_Name = "modTaskbar"
'* Taskbar stuff
Private picTaskIcon         As PictureBox
Public windowCount          As Long
Private TaskButtons(100)    As typTaskButton
Private Type typTaskButton
    Left    As Long
    Top     As Long
    Right   As Long
    Bottom  As Long
    Title   As String
    hWnd    As Long
End Type

Public downWhich    As Long
Public activeWhich  As Long
Public overWhich    As Long

Private Const bStretchButtons = False
Private Const xStart = 10
Private Const yStart = 1
Private Const XBuffer = 3
Private Const YBuffer = 3
Private Const ICON_SIZE = 16

'* Colors
Private clrButtonOff As Long  'GetSysColor(COLOR_3DFACE)
Private clrBorderOff As Long  'GetSysColor(COLOR_3DFACE)
Private Const clrTextOff = vbBlack
Private Const clrButtonOver = &HD2BDB6
Private Const clrBorderOver = &H6A240A
Private Const clrTextOver = vbBlack
Private Const clrButtonDown = &HB59285
Private Const clrBorderDown = &H6A240A
Private Const clrTextDown = vbBlack
Private Const clrButtonActive = &HD8D5D4
Private Const clrBorderActive = &H6A240A
Private Const clrTextActive = vbBlack
Private Const clrLines = &HA0A0A0
Private clrBackground As Long  'GetSysColor(COLOR_3DFACE)

'* API Stuff
Private Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type BitmapStruc
    hDcMemory As Long
    hDcBitmap As Long
    hDcPointer As Long
    Area As Rect
End Type

Private Type Size
    cx As Long
    cy As Long
End Type

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal i As Long, ByVal u As Long, ByVal S As Long, ByVal C As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal F As String) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal hIcon As Long) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As Size) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

Private Const COLOR_3DFACE As Long = 15
Private Const DI_MASK = &H1
Private Const DI_IMAGE = &H2
Private Const DI_NORMAL = DI_MASK Or DI_IMAGE
Private Const SRCCOPY = &HCC0020
Private Const WM_SETFOCUS = &H7

Public Function GetItemFromXY(X As Long, Y As Long) As Long
    Dim i As Integer
    For i = 1 To windowCount
        With TaskButtons(i)
            If X >= .Left And X <= .Right And Y >= .Top And Y <= .Bottom Then
                GetItemFromXY = i
                Exit Function
            End If
        End With
    Next i
    GetItemFromXY = -1
End Function

Function GetTaskText(picBox As PictureBox, strText As String, givenWidth As Integer) As String
    If picBox.TextWidth(strText) < givenWidth Then
        GetTaskText = strText
    Else
        Dim i As Integer
        For i = 1 To Len(strText)
            If picBox.TextWidth(Left(strText, i) & "...") > givenWidth - 5 Then
                GetTaskText = Left(strText, i - 1) & "..."
                Exit Function
            End If
        Next i
    End If
End Function
Sub DrawTaskbar(picBox As PictureBox)
    
    '* DrawTaskbar - redone 12/9/01 with API (again redone on 1/1/02)
    If windowCount = 0 Then Exit Sub
        
    Dim j As Integer, realWidth As Long, iconBuffer As Long, intButtonWidth As Long
    Dim strTitle As String, intWidth As Integer, i As Integer, iconY  As Long, strText As String
    Dim intStartY As Integer, intDown As Integer, TextX As Long, TextY As Long, TextHeight As Long
    Dim ly As Integer, lx As Integer, curX As Long, hFont2 As Long, TextWidth2 As Long
    Dim tBrush As Long, BMP As BitmapStruc, hFont As Long, theSize As Size  '* stuff for DC (buffer)
        
    '* Determine width of buttons
    If bStretchButtons Then
        realWidth = (picBox.ScaleWidth) - xStart - 1
    Else
        intButtonWidth = 125    'Change if you want
                                'I SAID CHANGE IT IF YOU WANT!
        realWidth = (windowCount) * intButtonWidth - xStart
        minSize = realWidth
    End If
    If realWidth + xStart + 2 >= picBox.ScaleWidth Then realWidth = picBox.ScaleWidth - xStart - 2
    
    minSize = (windowCount) * (ICON_SIZE + 40)
    If realWidth < minSize Then realWidth = minSize Else minSize = realWidth
    
    '* Create the fonts to be used
    hFont = CreateFont(13, 5, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "Tahoma")
    If hFont = 0 Then Exit Sub
    hFont2 = CreateFont(13, 6, 0, 0, 700, 0, 0, 0, 0, 0, 0, 0, 0, "Tahoma")
    If hFont2 = 0 Then Exit Sub
    
    '* Set the Area
    BMP.Area.Left = 0
    BMP.Area.Top = 0
    BMP.Area.Right = realWidth + xStart - 2
    BMP.Area.Bottom = picBox.ScaleHeight
    
    '* Create bitmap
    BMP.hDcMemory = CreateCompatibleDC(picBox.hdc)
    BMP.hDcBitmap = CreateCompatibleBitmap(picBox.hdc, picBox.ScaleWidth, picBox.ScaleHeight)
    BMP.hDcPointer = SelectObject(BMP.hDcMemory, BMP.hDcBitmap)
            
    If BMP.hDcMemory = 0 Or BMP.hDcBitmap = 0 Then
        DeleteObject BMP.hDcBitmap
        DeleteDC BMP.hDcMemory
        DeleteObject hFont
        Exit Sub
    End If
    
    '* Copy the background of picMenu into the DC
    tBrush = CreateSolidBrush(clrBackground)
    SelectObject BMP.hDcMemory, tBrush
    tPen = CreatePen(0, 0, clrBackground)
    SelectObject BMP.hDcMemory, tPen
    Rectangle BMP.hDcMemory, 0, 0, picBox.ScaleWidth + 1, picBox.ScaleHeight + 1
    DeleteObject tBrush
    DeleteObject tPen
   
    '* Set The Font
    Call SelectObject(BMP.hDcMemory, hFont)
    
    '* Draw the uh..thing on the left
    tBrush = CreateSolidBrush(clrLines)
    SelectObject BMP.hDcMemory, tBrush
    tPen = CreatePen(0, 1, clrLines)
    SelectObject BMP.hDcMemory, tPen
    lx = 3
    For ly = 6 To 18 Step 2
        Rectangle BMP.hDcMemory, 4, ly, 4 + lx, ly + 1
    Next ly
    DeleteObject tBrush
    DeleteObject tPen
    
    '* background of text transparent
    SetBkMode BMP.hDcMemory, 0
    
    intStartY = 1
    bReDraw = True
    
    If windowCount <= 0 Then GoTo finishit
    
    '* Set some variables
    intWidth = Int((realWidth / (windowCount)) - 0.5)
    GetTextExtentPoint32 BMP.hDcMemory, "wYz", 3, theSize
    TextHeight = theSize.cy
    TextY = (picBox.ScaleHeight - TextHeight) \ 2
    iconY = (picBox.ScaleHeight - ICON_SIZE) \ 2
    iconBuffer = ICON_SIZE + (XBuffer * 2)
    
    Dim child As Form
    i = 0
    
    curX = xStart
    
    For Each child In Forms
        On Error Resume Next
        If TypeOf child Is MDIForm Then GoTo nextem
        
        If downWhich - 1 = i Then
            SetTextColor BMP.hDcMemory, clrTextDown
            tBrush = CreateSolidBrush(clrButtonDown)
            SelectObject BMP.hDcMemory, tBrush
            tPen = CreatePen(0, 1, clrBorderDown)
            SelectObject BMP.hDcMemory, tPen
        ElseIf overWhich - 1 = i Then
            SetTextColor BMP.hDcMemory, clrTextOver
            tBrush = CreateSolidBrush(clrButtonOver)
            SelectObject BMP.hDcMemory, tBrush
            tPen = CreatePen(0, 1, clrBorderOver)
            SelectObject BMP.hDcMemory, tPen
            
            picBox.ToolTipText = " " & child.Caption & " "
        ElseIf activeWhich - 1 = i Then
            SetTextColor BMP.hDcMemory, clrTextActive
            tBrush = CreateSolidBrush(clrButtonActive)
            SelectObject BMP.hDcMemory, tBrush
            tPen = CreatePen(0, 1, clrBorderActive)
            SelectObject BMP.hDcMemory, tPen
        Else
            SetTextColor BMP.hDcMemory, clrTextOff
            tBrush = CreateSolidBrush(clrButtonOff)
            SelectObject BMP.hDcMemory, tBrush
            tPen = CreatePen(0, 1, clrBorderOff)
            SelectObject BMP.hDcMemory, tPen
        End If
        Rectangle BMP.hDcMemory, curX, yStart + 1, curX + intWidth - 1, picBox.ScaleHeight 'YStart + intHeight
        DeleteObject tBrush
        DeleteObject tPen
        
        If (downWhich - 1 = i) Then
            intDown = 1
        Else
            intDown = 0
        End If
        
        DrawIconEx BMP.hDcMemory, curX + XBuffer * 2 + intDown + 1, iconY + intDown, child.Icon, ICON_SIZE, ICON_SIZE, 0, 0, DI_NORMAL
                
        '* draw actual text
        strText = GetTaskText(picBox, child.Caption, intWidth - ICON_SIZE - (XBuffer * 2) - 6)
        TextOut BMP.hDcMemory, curX + XBuffer * 2 + intDown + iconBuffer, TextY + intDown, strText, Len(strText)
        
        With TaskButtons(i + 1)
            .Title = child.Caption
            .Left = curX
            .Right = curX + intWidth - 1
            .Top = yStart + 1
            .Bottom = picBox.ScaleHeight
            .hWnd = child.hWnd
        End With
        
        '* increment variables
        i = i + 1
        curX = curX + intWidth
        
nextem:
    Next child
    
finishit:
    BitBlt picBox.hdc, BMP.Area.Left, BMP.Area.Top, BMP.Area.Right, BMP.Area.Bottom, BMP.hDcMemory, 0, 0, SRCCOPY
    
    DeleteObject tBrush
    DeleteObject tPen
    DeleteObject hFont2
    DeleteObject hFont
    DeleteObject BMP.hDcBitmap
    DeleteDC BMP.hDcMemory
    bDrew = True
End Sub


Public Sub InitTaskbar()
    clrButtonOff = GetSysColor(COLOR_3DFACE)
    clrBorderOff = GetSysColor(COLOR_3DFACE)
    clrBackground = GetSysColor(COLOR_3DFACE)
    'set pictaskicon = new
    'MsgBox frmChild.Icon
    
End Sub

Sub SetWindowCount()
    windowCount = 0
    Dim child As Form
    For Each child In Forms
        If child.Visible And Not TypeOf child Is MDIForm Then
            windowCount = windowCount + 1
        End If
    Next child
End Sub


Public Sub SetFocusByIndex(which As Long)
    If which = -1 Then Exit Sub
    
    SendMessage TaskButtons(which).hWnd, WM_SETFOCUS, 0, 0
End Sub


Public Sub SetFocusByHwnd(picBox As PictureBox, hWnd As Long)
    Dim i As Integer
    For i = 1 To windowCount
        If hWnd = TaskButtons(i).hWnd Then
            activeWhich = i
            Exit For
        End If
    Next i
            
    DrawTaskbar picBox
End Sub
