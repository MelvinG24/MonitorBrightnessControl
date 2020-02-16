Attribute VB_Name = "mdlRoundFrm"
Public xp As Long, yp As Long
Public mShape As Integer
Public mChildFormRegion As Long

Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Public Sub frmRounded(frm As Form)
    mShape = 1
    xp = Screen.TwipsPerPixelX
    yp = Screen.TwipsPerPixelY
      
    If mShape = 1 Then
        mChildFormRegion = CreateRoundRectRgn(0, 0, frm.Width / xp, frm.Height / yp, 12, 12)
    Else
        mChildFormRegion = CreateEllipticRgn(0, 0, frm.Width / xp, frm.Height / yp)
    End If
    
    SetWindowRgn frm.hwnd, mChildFormRegion, False
End Sub
