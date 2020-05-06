Attribute VB_Name = "basFrom"
Option Explicit

Private Declare Function apiSetWindowPos Lib "user32" Alias "SetWindowPos" ( _
        ByVal hwnd As Long, _
        ByVal hWndInsertAfter As Long, _
        ByVal X As Long, _
        ByVal Y As Long, _
        ByVal cx As Long, _
        ByVal cy As Long, _
        ByVal wFlags As Long) As Long

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
        apiSetWindowPos frm.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, SWP_NOMOVE + SWP_NOSIZE
    Else
        apiSetWindowPos frm.hwnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, SWP_NOMOVE + SWP_NOSIZE
    End If
End Sub
