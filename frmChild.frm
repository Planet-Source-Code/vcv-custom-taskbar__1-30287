VERSION 5.00
Begin VB.Form frmChild 
   Caption         =   "Child Window"
   ClientHeight    =   2175
   ClientLeft      =   4995
   ClientTop       =   4440
   ClientWidth     =   4305
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2175
   ScaleWidth      =   4305
End
Attribute VB_Name = "frmChild"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    SetFocusByHwnd MDI.picTaskbar, Me.hWnd
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SetWindowCount
    DrawTaskbar MDI.picTaskbar
End Sub


