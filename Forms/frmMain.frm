VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   90
   ClientTop       =   13020
   ClientWidth     =   4560
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   2280
      Top             =   0
   End
   Begin VB.Timer Timer2 
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

Dim Repeticiones As Integer
Dim activeWin As Long
Dim retVal As Long

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

Private Declare Function GetTopWindow Lib "user32" ( _
                ByVal hwnd As Long) As Long

Private Const GWL_EXSTYLE = (-20)
Private Const GWL_STYLE = (-16)
Private Const HWND_NOTOPMOST = -2
Private Const HWND_TOPMOST = -1
Private Const LWA_ALPHA = &H2
Private Const LWA_COLORKEY = &H1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Private Const WS_EX_LAYERED = &H80000
Private Const WS_EX_TRANSPARENT = &H20&

Private Sub SetWinToTOP()
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Repeticiones = Repeticiones + 1
        If Timer3.Enabled = False Then Timer3.Enabled = True
        If Repeticiones >= 3 Then
            P_VarBrightnessLevel = 128
            OnOffSwitch 2
        End If
    Else
        Repeticiones = 0
    End If
End Sub

Private Sub Form_Load()
    'Set window color
    Me.BackColor = vbBlack
    
    'Set window transparency
    SetWindowLong Me.hwnd, GWL_EXSTYLE, WS_EX_LAYERED
    
    'Set window click-through enable
    SetWindowLong Me.hwnd, GWL_EXSTYLE, GetWindowLong(Me.hwnd, GWL_EXSTYLE) Or WS_EX_TRANSPARENT
    
    'Set window transparency percentage
    SetLayeredWindowAttributes Me.hwnd, vbBlack, P_VarBrightnessLevel, LWA_ALPHA
    
    'Set main windows position to the top
    SetWinToTOP
End Sub

Private Sub Form_Resize()
    'Set Repeticiones to 0
    Repeticiones = 0
    
    'Set shortcut label
    lblInfo.Visible = P_VarChckSCEnable
    lblInfo.Caption = LoadResString(101) + " " + P_VarRsBrightness & vbNewLine & LoadResString(102) + " " + P_VarLwBrightness
    
    'Set label position
    lblInfo.Top = (WindowRect.Bottom * Screen.TwipsPerPixelY - lblInfo.Height) - 120
    lblInfo.Left = (WindowRect.Right * Screen.TwipsPerPixelX - lblInfo.Width) - 1320
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Timer1.Enabled = False
    Timer2.Enabled = False
    Timer3.Enabled = False
End Sub

Private Sub Timer1_Timer()
    SetLayeredWindowAttributes Me.hwnd, vbBlack, P_VarBrightnessLevel, LWA_ALPHA
End Sub

Private Sub Timer2_Timer()
    activeWin = GetTopWindow(Me.hwnd)
    If activeWin <> Me.hwnd Then
        SetWinToTOP
    End If
End Sub

Private Sub Timer3_Timer()
    If Not Repeticiones >= 3 Then
        Repeticiones = 0
        Timer3.Enabled = False
    End If
End Sub
