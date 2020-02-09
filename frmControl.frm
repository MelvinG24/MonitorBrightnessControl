VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmControl 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4185
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   975
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   975
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.Slider sliderControl 
      Height          =   3255
      Left            =   200
      TabIndex        =   2
      Top             =   120
      Width           =   600
      _ExtentX        =   1058
      _ExtentY        =   5741
      _Version        =   393216
      OLEDropMode     =   1
      Orientation     =   1
      Max             =   100
      TickStyle       =   2
      TickFrequency   =   5
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   975
      TabIndex        =   0
      Top             =   3690
      Width           =   975
      Begin VB.Label lblConfig 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Config"
         BeginProperty Font 
            Name            =   "Ubuntu Mono"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   185
         TabIndex        =   1
         Top             =   120
         Width           =   630
      End
   End
End
Attribute VB_Name = "frmControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" ( _
                ByVal uAction As Long, _
                ByVal uParam As Long, _
                ByRef lpvParam As Any, _
                ByVal fuWinIni As Long) As Long

Private Const SPI_GETWORKAREA = 48

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Sub PositionForm()
    
End Sub

Private Sub Form_Load()
    Dim WindowRect As RECT
    SystemParametersInfo SPI_GETWORKAREA, 0, WindowRect, 0
    Me.Top = WindowRect.Bottom * Screen.TwipsPerPixelY - Me.Height
    Me.Left = WindowRect.Right * Screen.TwipsPerPixelX - Me.Width
End Sub

Private Sub Form_LostFocus()
    Unload Me
End Sub
