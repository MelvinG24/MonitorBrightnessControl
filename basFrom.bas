Attribute VB_Name = "basFrom"
Option Explicit

Private Declare Function apiSetWindowPos Lib "user32" Alias "SetWindowPos" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const HWND_TOPMOST As Long = -1
Private Const HWND_NOTOPMOST As Long = -2
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOSIZE As Long = &H1

Private Sub Main()
    If App.PrevInstance Then Exit Sub
    Load frmSysTray
End Sub

Public Sub SetOnTop(frm As Form, OnTop As Long)
    If OnTop = -1 Then
        apiSetWindowPos frm.hWnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, SWP_NOMOVE + SWP_NOSIZE
    Else
        apiSetWindowPos frm.hWnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, SWP_NOMOVE + SWP_NOSIZE
    End If
End Sub
