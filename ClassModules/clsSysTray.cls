VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSysTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Written by Ellis Dee of vbforums.com
'
' Thanks to...
'   Edgemeal   http://www.vbforums.com/showthread.php?t=595193
'   Mike D. Sutton   http://www.mvps.org/EDais/
'   Steve McMahon   http://www.vbaccelerator.com/home/VB/Code/Libraries/Shell_Projects/SysTray_-_The_Easy_Way/article.asp
'   Ben Baird for the original code
Option Explicit

' Tray Icon
Private Const NIF_ICON = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_TIP = &H4
Private Const NIF_STATE = &H8
Private Const NIF_INFO = &H10

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIM_SETFOCUS = &H3
Private Const NIM_SETVERSION = &H4

Private Const NOTIFYICON_VERSION = 3

Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 128
    dwState As Long
    dwStateMask As Long
    szInfo As String * 256
    uTimeoutOrVersion As Long
    szInfoTitle As String * 64
    dwInfoFlags As Long
End Type
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

' Bitmap to Icon
Private Type IconInfo
    fIcon As Long
    xHotspot As Long
    yHotspot As Long
    hbmMask As Long
    hbmColor As Long
End Type
Private Type Bitmap
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type
Private Declare Function BitBlt Lib "gdi32" (ByVal hdcDest As Long, ByVal nXDest As Long, ByVal nYDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hdcSrc As Long, ByVal nXSrc As Long, ByVal nYSrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, ByRef lpBits As Any) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateIconIndirect Lib "user32" (ByRef pIconInfo As IconInfo) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long

' Tray Events (fire in frmSysTray.frm)
Public Event LeftClick()
Public Event RightClick()
Public Event DoubleClick()
Public Event BalloonHide()
Public Event BalloonTimeOut()
Public Event BalloonClicked()
Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Const WM_USER = &H400
Private Const NIN_SELECT = WM_USER
Private Const NINF_KEY = &H1
Private Const NIN_KEYSELECT = (NIN_SELECT Or NINF_KEY)
Private Const NIN_BALLOONSHOW = (WM_USER + 2)
Private Const NIN_BALLOONHIDE = (WM_USER + 3)
Private Const NIN_BALLOONTIMEOUT = (WM_USER + 4)
Private Const NIN_BALLOONUSERCLICK = (WM_USER + 5)

' Ballon tips
Public Enum BalloonIconEnum
   beNone = 0
   beInformation = 1
   beWarning = 2
   beError = 3
   beNoSound = &H10
End Enum

Private mstrTooltip As String

Private mtypIcon As NOTIFYICONDATA


' METHODS


' Set tray icon to contents of 32x32 picturebox
Public Sub DrawIcon(ppic As PictureBox, Optional plngTransparentColor As Long = -1)
    ppic.AutoRedraw = True
    ppic.Picture = ppic.Image
    If plngTransparentColor < 0 Then
        mtypIcon.hIcon = BitmapToIcon(ppic.Picture.Handle)
    Else
        mtypIcon.hIcon = BitmapToIconTransparent(ppic.Picture.Handle, plngTransparentColor)
    End If
    If mtypIcon.hIcon Then
        RefreshIcon
        DestroyIcon mtypIcon.hIcon
    End If
End Sub

Public Function Init(pfrm As Form, pstrTooltip As String) As Boolean
    Const NIF_MESSAGE = &H1
    Const WM_MOUSEMOVE = &H200
    Const NIM_ADD = &H0
    
    With mtypIcon
        .cbSize = Len(mtypIcon)
        .hwnd = pfrm.hwnd
        .hIcon = pfrm.Icon
        .uID = vbNull
        .uFlags = NIF_MESSAGE
        .uCallbackMessage = WM_MOUSEMOVE
    End With
    Shell_NotifyIcon NIM_ADD, mtypIcon
    SetIcon pfrm
    Me.TooltipText = pstrTooltip
End Function

Public Sub MouseMove(Button As Integer, ByVal X As Long, pfrm As Form)
   Select Case X
        Case WM_LBUTTONDBLCLK: RaiseEvent DoubleClick
        Case WM_LBUTTONUP: RaiseEvent LeftClick
        Case WM_RBUTTONUP: RaiseEvent RightClick
        Case NIN_BALLOONHIDE: RaiseEvent BalloonHide
        Case NIN_BALLOONTIMEOUT: RaiseEvent BalloonTimeOut
        Case NIN_BALLOONUSERCLICK: RaiseEvent BalloonClicked
   End Select
End Sub

' Set tray icon to a form icon
Public Sub SetIcon(pfrm As Form)
    mtypIcon.hIcon = pfrm.Icon
    RefreshIcon
End Sub

Public Sub ShowBalloonTip(ByVal pstrMessage As String, Optional ByVal penIcon As BalloonIconEnum, Optional ByVal pstrTitle As String, Optional ByVal plngTimeout = 30000)
    mtypIcon.szInfo = pstrMessage & vbNullChar
    mtypIcon.szInfoTitle = pstrTitle & vbNullChar
    mtypIcon.uTimeoutOrVersion = plngTimeout
    mtypIcon.dwInfoFlags = penIcon
    mtypIcon.uFlags = NIF_INFO
    Shell_NotifyIcon NIM_MODIFY, mtypIcon
End Sub


' PROPERTIES


Public Property Get TooltipText() As String
    TooltipText = mstrTooltip
End Property

Public Property Let TooltipText(ByVal pstrTooltip As String)
    Const NIF_TIP = &H4
    Const NIM_MODIFY = &H1
    
    mstrTooltip = pstrTooltip
    mtypIcon.szTip = mstrTooltip & vbNullChar
    mtypIcon.uFlags = NIF_TIP
    Shell_NotifyIcon NIM_MODIFY, mtypIcon
End Property


' INTERNAL ROUTINES


Private Sub Class_Terminate()
    Const NIM_DELETE = &H2
    
    Shell_NotifyIcon NIM_DELETE, mtypIcon
    Set frmSysTray = Nothing
End Sub

Private Sub RefreshIcon()
    Const NIF_ICON = &H2
    Const NIM_MODIFY = &H1
    
    mtypIcon.uFlags = NIF_ICON
    Shell_NotifyIcon NIM_MODIFY, mtypIcon
End Sub

' Birmap to icon functions provided by Edgemeal
' Special Thanks to Mike D. Sutton - Http://www.mvps.org/EDais/
Private Function BitmapToIcon(ByVal inBMP As Long) As Long
    Dim IconInf As IconInfo
    Dim BMInf As Bitmap
    Dim hMask As Long
    
    ' Get some information about this Bitmap and create a mask the same size
    If (GetObject(inBMP, Len(BMInf), BMInf) = 0) Then Exit Function
    hMask = CreateBitmap(BMInf.bmWidth, BMInf.bmHeight, 0, 0, ByVal 0&)
    With IconInf ' Set some information about the icon
        .fIcon = True
        .hbmMask = hMask
        .hbmColor = inBMP
    End With
    ' Create the icon and destroy the temp mask
    BitmapToIcon = CreateIconIndirect(IconInf)
    Call DeleteObject(hMask)
End Function

' Take a HBITMAP and return an HICON
' Modified by Edgemeal for 32x32 pixel type Tray Icon programs.
Private Function BitmapToIconTransparent(ByVal hSrcBMP As Long, Optional ByVal inTransCol As Long = -1) As Long
    Dim IconInf As IconInfo
    Dim hSrcDC As Long ', hSrcBMP As Long
    Dim hSrcOldBMP As Long
    Dim hMaskDC As Long
    Dim hMaskBMP As Long
    Dim hMaskOldBMP As Long
    
    ' Create DC's and select source copy
    hSrcDC = CreateCompatibleDC(0)
    hMaskDC = CreateCompatibleDC(0)
    hSrcOldBMP = SelectObject(hSrcDC, hSrcBMP)
    If (hSrcOldBMP) Then ' Extract a colour mask from source copy
        hMaskBMP = GetColMask(hSrcDC, inTransCol)
        hMaskOldBMP = SelectObject(hMaskDC, hMaskBMP)
        If (hMaskOldBMP) Then ' Overlay inverted mask over source
            Call SetTextColor(hSrcDC, vbWhite)
            Call SetBkColor(hSrcDC, vbBlack)
            Call BitBlt(hSrcDC, 0, 0, 32, 32, hMaskDC, 0, 0, vbSrcAnd)
            Call SelectObject(hMaskDC, hMaskOldBMP) ' De-select mask
        End If
        ' De-select source copy
        Call SelectObject(hSrcDC, hSrcOldBMP)
    End If
    ' Destroy DC's
    Call DeleteDC(hMaskDC)
    Call DeleteDC(hSrcDC)
    With IconInf ' Set some information about the icon
        .fIcon = True
        .hbmMask = hMaskBMP
        .hbmColor = hSrcBMP
    End With
    ' Create the icon and destroy the temp mask
    BitmapToIconTransparent = CreateIconIndirect(IconInf)
    ' Destroy interim Bitmaps
    Call DeleteObject(hMaskBMP)
    Call DeleteObject(hSrcBMP)
End Function

Private Function GetColMask(ByVal inDC As Long, ByVal inMaskCol As Long) As Long
    Dim MaskDC As Long, MaskBMP As Long, OldMask As Long, OldBack As Long
    ' Create a new DC
    MaskDC = CreateCompatibleDC(inDC)
    If (MaskDC) Then ' Create a new 1-bpp Bitmap (DDB)
        MaskBMP = CreateBitmap(32, 32, 1, 1, ByVal 0&)
        If (MaskBMP) Then ' Select Bitmap into DC
            OldMask = SelectObject(MaskDC, MaskBMP)
            If (OldMask) Then ' Set mask colour
                OldBack = SetBkColor(inDC, inMaskCol)
                ' Generate mask image
                If (BitBlt(MaskDC, 0, 0, 32, 32, inDC, 0, 0, vbSrcCopy) <> 0) Then GetColMask = MaskBMP
                ' Clean up
                Call SetBkColor(inDC, OldBack)
                Call SelectObject(MaskDC, OldMask)
            End If
            ' Something went wrong, destroy mask Bitmap
            If (GetColMask = 0) Then Call DeleteObject(MaskBMP)
        End If
        ' Destroy temporary DC
        Call DeleteDC(MaskDC)
    End If
End Function
