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
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1800
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
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
                ByVal hwnd As Long, _
                ByVal nIndex As Long) As Long
 
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
                ByVal hwnd As Long, _
                ByVal nIndex As Long, _
                ByVal dwNewLong As Long) As Long
                
Private Declare Function SetLayeredWindowAttributes Lib "user32" ( _
                ByVal hwnd As Long, _
                ByVal crKey As Long, _
                ByVal bAlpha As Byte, _
                ByVal dwFlags As Long) As Long

Private Declare Function SetWindowPos Lib "user32" ( _
                ByVal hwnd As Long, _
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

Private Const WS_EX_COMPOSITED = &H2
Private Const SPI_GETWORKAREA = 48
Private Const GWL_STYLE = (-16)
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const WS_EX_TRANSPARENT = &H20&         'new
Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOSIZE As Long = &H1
Private Const FLAGS = SWP_NOMOVE + SWP_NOSIZE
Private Const HWND_TOPMOST = -1                 'new
Private Const HWND_NOTOPMOST = -2
Attribute HWND_NOTOPMOST.VB_VarHelpID = -1

Private Sub ClickThru(frm As Form, bEnabled As Boolean)
    If bEnabled = True Then ' enable click-thru form
        SetWindowLong frm.hwnd, GWL_EXSTYLE, GetWindowLong(frm.hwnd, GWL_EXSTYLE) Or WS_EX_TRANSPARENT
    Else ' disable click thru
        SetWindowLong frm.hwnd, GWL_EXSTYLE, GetWindowLong(frm.hwnd, GWL_EXSTYLE) And Not WS_EX_TRANSPARENT
    End If
End Sub

Private Sub Form_Load()
    'Set values to the variable
    frmSysTray.STATE_SCREEN = True
    SystemParametersInfo SPI_GETWORKAREA, 0, WindowRect, 0
    
    'Shortcut label
    lblInfo.Visible = frmSysTray.STATE_SCREEN
    lblInfo.Caption = "Lower-Brightness: " + frmSysTray.tb & vbNewLine & "Raise-Brightness: " + frmSysTray.ts
    
    'Black screen settings
    Me.BackColor = vbBlack
    
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS 'Set formulario always on top
    SetWindowLong Me.hwnd, GWL_EXSTYLE, WS_EX_LAYERED
    'SetWindowLong Me.hwnd, GWL_EXSTYLE, GetWindowLong(Me.hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED
    SetLayeredWindowAttributes Me.hwnd, vbBlack, frmSysTray.M_BRIGHTNESS, LWA_ALPHA
    
    ClickThru Me, True 'Enable click-thru formulario
    If Me.Visible = True Then
        Timer1.Enabled = True
        Timer2.Enabled = True
    End If
End Sub

Private Sub Form_Paint()
    SetOnTop Me, True
    lblInfo.Visible = frmSysTray.STATE_SCREEN
    SetLayeredWindowAttributes Me.hwnd, vbBlack, frmSysTray.M_BRIGHTNESS, LWA_ALPHA
End Sub

Private Sub Form_Resize()
    lblInfo.Top = (WindowRect.Bottom * Screen.TwipsPerPixelY - lblInfo.Height) - 120
    lblInfo.Left = (WindowRect.Right * Screen.TwipsPerPixelX - lblInfo.Width) - 800
End Sub

Private Sub sliderControl_KeyDown(KeyCode As Integer, Shift As Integer)
    'If KeyCode = 116 Then
    '    sliderControl.Value = sliderControl.Value - 10
    'ElseIf KeyCode = 117 Then
    '    sliderControl.Value = sliderControl.Value + 10
    'End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmSysTray.STATE_SCREEN = False
    Timer1.Enabled = False
    Timer2.Enabled = False
End Sub

