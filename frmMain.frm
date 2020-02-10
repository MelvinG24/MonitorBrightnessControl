VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4560
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1320
      Top             =   0
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Ubuntu Mono"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000010&
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   810
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WindowRect As RECT

Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" ( _
                ByVal uAction As Long, _
                ByVal uParam As Long, _
                ByRef lpvParam As Any, _
                ByVal fuWinIni As Long) As Long

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
                ByVal hWnd As Long, _
                ByVal nIndex As Long) As Long
 
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
                ByVal hWnd As Long, _
                ByVal nIndex As Long, _
                ByVal dwNewLong As Long) As Long
                
Private Declare Function SetLayeredWindowAttributes Lib "user32" ( _
                ByVal hWnd As Long, _
                ByVal crKey As Long, _
                ByVal bAlpha As Byte, _
                ByVal dwFlags As Long) As Long

Private Declare Function SetWindowPos Lib "user32" ( _
                ByVal hWnd As Long, _
                ByVal hWndInsertAfter As Long, _
                ByVal X As Long, _
                ByVal Y As Long, _
                ByVal cx As Long, _
                ByVal cy As Long, _
                ByVal wFlags As Long) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Const SPI_GETWORKAREA = 48
Private Const GWL_STYLE = (-16)
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const WS_EX_TRANSPARENT = &H20&         'new
Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2
Private Const SWP_NOMOVE = 2                    'new
Private Const SWP_NOSIZE = 1                    'new
Private Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE  'new
Private Const HWND_TOPMOST = -1                 'new
Private Const HWND_NOTOPMOST = -2
Attribute HWND_NOTOPMOST.VB_VarHelpID = -1

Private Sub ClickThru(Frm As Form, bEnabled As Boolean)
    If bEnabled = True Then ' enable click-thru form
        SetWindowLong Frm.hWnd, GWL_EXSTYLE, GetWindowLong(Frm.hWnd, GWL_EXSTYLE) Or WS_EX_TRANSPARENT
    Else ' disable click thru
        SetWindowLong Frm.hWnd, GWL_EXSTYLE, GetWindowLong(Frm.hWnd, GWL_EXSTYLE) And Not WS_EX_TRANSPARENT
    End If
End Sub

Private Sub Form_Load()
    'Set values to the variable
    SystemParametersInfo SPI_GETWORKAREA, 0, WindowRect, 0
    
    'Shortcut label
    lblInfo.Visible = frmSysTray.L_SHORTCUTS
    lblInfo.Caption = "Lower-Brightness: " + frmSysTray.tb & vbNewLine & "Raise-Brightness: " + frmSysTray.ts
    lblInfo.Top = WindowRect.Bottom * Screen.TwipsPerPixelY - lblInfo.Height
    lblInfo.Left = WindowRect.Right * Screen.TwipsPerPixelX - lblInfo.Width
    
    'Black screen settings
    Me.BackColor = vbBlack
    
    SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS 'Set formulario always on top
    SetWindowLong Me.hWnd, GWL_EXSTYLE, WS_EX_LAYERED
    'SetWindowLong Me.hwnd, GWL_EXSTYLE, GetWindowLong(Me.hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED
    SetLayeredWindowAttributes Me.hWnd, vbBlack, frmSysTray.M_BRIGHTNESS, LWA_ALPHA
    
    ClickThru Me, True 'Enable click-thru formulario
End Sub

Private Sub Form_Resize()
    lblInfo.Top = (WindowRect.Bottom * Screen.TwipsPerPixelY - lblInfo.Height) - 120
    lblInfo.Left = (WindowRect.Right * Screen.TwipsPerPixelX - lblInfo.Width) - 120
End Sub

Private Sub sliderControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        sliderControl.Value = sliderControl.Value - 10
    ElseIf KeyCode = 117 Then
        sliderControl.Value = sliderControl.Value + 10
    End If
End Sub

Private Sub Timer1_Timer()
    lblInfo.Visible = frmSysTray.L_SHORTCUTS
    SetLayeredWindowAttributes Me.hWnd, vbBlack, frmSysTray.M_BRIGHTNESS, LWA_ALPHA
End Sub
