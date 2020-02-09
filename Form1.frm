VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame frameControl 
      Caption         =   "1255"
      Height          =   975
      Left            =   1440
      TabIndex        =   0
      Top             =   2640
      Width           =   4695
      Begin MSComctlLib.Slider sliderControl 
         Height          =   435
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   767
         _Version        =   393216
         Max             =   100
         SelStart        =   50
         TickStyle       =   3
         Value           =   50
      End
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Teclado:"
      BeginProperty Font 
         Name            =   "Ubuntu Mono"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   1680
      TabIndex        =   2
      Top             =   1920
      Width           =   1080
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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
                ByVal x As Long, _
                ByVal y As Long, _
                ByVal cx As Long, _
                ByVal cy As Long, _
                ByVal wFlags As Long) As Long

 
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

Private Sub ClickThru(Frm As Form, bEnabled As Boolean)
    If bEnabled = True Then ' enable click-thru form
        SetWindowLong Frm.hwnd, GWL_EXSTYLE, GetWindowLong(Frm.hwnd, GWL_EXSTYLE) Or WS_EX_TRANSPARENT
    Else ' disable click thru
        SetWindowLong Frm.hwnd, GWL_EXSTYLE, GetWindowLong(Frm.hwnd, GWL_EXSTYLE) And Not WS_EX_TRANSPARENT
    End If
End Sub

Private Sub Form_Load()
    Dim tb As String
    Dim ts As String
    Dim r As Integer
    
    tb = "F5"
    ts = "F6"
    r = (sliderControl.Value * 25.5) / 10
    
    lblInfo.Caption = "Bajar brillo: " + tb & vbNewLine & "Subir brillo: " + ts
    frameControl.Caption = r
    
    frameControl.Left = (Me.Width - frameControl.Width) / 2
    frameControl.Top = (Me.Height - frameControl.Height) / 2
    lblInfo.Left = (Me.Width - lblInfo.Width) - 120
    lblInfo.Top = (Me.Height - lblInfo.Height) - 120
    
    Me.BackColor = vbBlack
    
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS 'Set formulario always on top
    SetWindowLong Me.hwnd, GWL_EXSTYLE, WS_EX_LAYERED
    'SetWindowLong Me.hwnd, GWL_EXSTYLE, GetWindowLong(Me.hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED
    SetLayeredWindowAttributes Me.hwnd, vbBlack, r, LWA_ALPHA
    
    ClickThru Me, True 'Enable click-thru formulario
End Sub

Private Sub Form_Resize()
    frameControl.Left = (Me.Width - frameControl.Width) / 2
    frameControl.Top = (Me.Height - frameControl.Height) / 2
    lblInfo.Left = (Me.Width - lblInfo.Width) - 120
    lblInfo.Top = (Me.Height - lblInfo.Height) - 120
End Sub

Private Sub sliderControl_Change()
    Dim r As Integer
    
    r = (sliderControl.Value * 25.5) / 10
    
    frameControl.Caption = r
    
    SetLayeredWindowAttributes Me.hwnd, vbBlack, r, LWA_ALPHA
End Sub

Private Sub sliderControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        sliderControl.Value = sliderControl.Value - 10
    ElseIf KeyCode = 117 Then
        sliderControl.Value = sliderControl.Value + 10
    End If
End Sub
