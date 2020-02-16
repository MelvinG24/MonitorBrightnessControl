Attribute VB_Name = "mdlrmvBG"
Option Explicit
 
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
                ByVal hwnd As Long, _
                ByVal nIndex As Long, _
                ByVal dwNewLong As Long) As Long
                
Private Declare Function SetLayeredWindowAttributes Lib "user32" ( _
                ByVal hwnd As Long, _
                ByVal crKey As Long, _
                ByVal bAlpha As Byte, _
                ByVal dwFlags As Long) As Long

Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_COLORKEY = &H1

Public Function removBG(frm As Form, color As Integer)
    frm.BackColor = color
    
    SetWindowLong frm.hwnd, GWL_EXSTYLE, WS_EX_LAYERED
    SetLayeredWindowAttributes frm.hwnd, color, 0&, LWA_COLORKEY
End Function


