VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MsComCtl.ocx"
Begin VB.Form frmControl 
   BorderStyle     =   0  'None
   ClientHeight    =   4335
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4095
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   4335
   ScaleMode       =   0  'User
   ScaleWidth      =   4410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MonitorBrightnessControl.ctrlSlider ctrlSlider1 
      Height          =   2175
      Left            =   105
      TabIndex        =   1
      Top             =   720
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   3836
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   480
      Top             =   3720
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   3720
   End
   Begin MSComctlLib.Slider sliderControl 
      Height          =   2295
      Left            =   1200
      TabIndex        =   0
      Top             =   720
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   4048
      _Version        =   393216
      OLEDropMode     =   1
      Orientation     =   1
      LargeChange     =   10
      SmallChange     =   10
      Max             =   100
      SelStart        =   50
      TickStyle       =   2
      TickFrequency   =   50
      Value           =   50
   End
   Begin VB.Image imgOnOff 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   93
      MouseIcon       =   "frmControl.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "frmControl.frx":0152
      Stretch         =   -1  'True
      Top             =   85
      Width           =   285
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      X1              =   100.154
      X2              =   399.538
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Image imgSettingsOn 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   102
      MouseIcon       =   "frmControl.frx":0E1F
      MousePointer    =   99  'Custom
      Picture         =   "frmControl.frx":0F71
      Stretch         =   -1  'True
      Top             =   3195
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image imgSettingsOff 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   102
      MouseIcon       =   "frmControl.frx":21E4
      MousePointer    =   99  'Custom
      Picture         =   "frmControl.frx":2336
      Stretch         =   -1  'True
      Top             =   3195
      Width           =   285
   End
   Begin VB.Image imgOn 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   600
      MouseIcon       =   "frmControl.frx":35A9
      MousePointer    =   99  'Custom
      Picture         =   "frmControl.frx":36FB
      Stretch         =   -1  'True
      Top             =   120
      Width           =   285
   End
   Begin VB.Image imgOff 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   600
      MouseIcon       =   "frmControl.frx":43C8
      MousePointer    =   99  'Custom
      Picture         =   "frmControl.frx":451A
      Stretch         =   -1  'True
      Top             =   480
      Width           =   285
   End
   Begin VB.Shape shapeCon 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   3000
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   585
      Width           =   465
   End
   Begin VB.Shape shapeOnOff 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   465
   End
   Begin VB.Image imgDay 
      Height          =   495
      Left            =   1320
      MouseIcon       =   "frmControl.frx":51E7
      MousePointer    =   99  'Custom
      Picture         =   "frmControl.frx":54F1
      Stretch         =   -1  'True
      Top             =   120
      Width           =   495
   End
   Begin VB.Image imgNight 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   1200
      MouseIcon       =   "frmControl.frx":57CDE
      MousePointer    =   99  'Custom
      Picture         =   "frmControl.frx":57FE8
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   495
   End
End
Attribute VB_Name = "frmControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Dim sys As StatusBar
Dim ActiveApp As Long
Dim r As Integer
Dim WindowRect As RECT

Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" ( _
                ByVal uAction As Long, _
                ByVal uParam As Long, _
                ByRef lpvParam As Any, _
                ByVal fuWinIni As Long) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Const SPI_GETWORKAREA = 48

Private Sub ctrlSlider1_Change()
    r = (ctrlSlider1.Value * 25.5) / 10
    frmSysTray.M_BRIGHTNESS = r
End Sub

Private Sub Form_GotFocus()
    ActiveApp = GetActiveWindow
End Sub

Private Sub Form_Load()
    removBG Me, vbRed
    
    
    ctrlSlider1.Value = sliderControl.Value
    
    ActiveApp = 0
    r = (frmSysTray.M_BRIGHTNESS * 10) / 25.5
    sliderControl.Value = r
    Timer1.Enabled = True
    SystemParametersInfo SPI_GETWORKAREA, 0, WindowRect, 0
    
    If frmSysTray.STATE_SCREEN = True Then
        imgOn.Visible = True
        imgOff.Visible = False
    ElseIf frmSysTray.STATE_SCREEN = False Then
        imgOn.Visible = False
        imgOff.Visible = True
    End If
End Sub

Private Sub form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If X >= Me.Left And X <= Me.Left + Me.Width And _
'        Y >= Me.Top And Y <= Me.Top + Me.Height Then
'        MsgBox "Shape1 has been clicked."
'    End If
End Sub

Private Sub Form_LostFocus()
    Unload Me
End Sub

Private Sub Form_Resize()
    Me.Width = shapeCon.Width
    Me.Height = shapeCon.Top + shapeCon.Height
    
    Me.Top = (WindowRect.Bottom * Screen.TwipsPerPixelY - Me.Height) - 120
    Me.Left = (WindowRect.Right * Screen.TwipsPerPixelX - Me.Width) - 120
End Sub

Private Sub imgDay_Click()
    If sliderControl.Value >= 76 And sliderControl.Value <= 100 Then
        sliderControl.Value = 75
    ElseIf sliderControl.Value >= 51 And sliderControl.Value <= 75 Then
        sliderControl.Value = 50
    ElseIf sliderControl.Value >= 26 And sliderControl.Value <= 50 Then
        sliderControl.Value = 25
    ElseIf sliderControl.Value <= 25 Then
        sliderControl.Value = 0
    End If
End Sub

Private Sub imgDay_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgDay.Appearance = 1
    imgDay.BorderStyle = 1
End Sub

Private Sub imgDay_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgDay.Appearance = 0
    imgDay.BorderStyle = 0
End Sub

Private Sub imgNight_Click()
    If sliderControl.Value >= 0 And sliderControl.Value <= 24 Then
        sliderControl.Value = 25
    ElseIf sliderControl.Value >= 25 And sliderControl.Value <= 49 Then
        sliderControl.Value = 50
    ElseIf sliderControl.Value >= 50 And sliderControl.Value <= 74 Then
        sliderControl.Value = 75
    ElseIf sliderControl.Value >= 75 Then
        sliderControl.Value = 100
    End If
End Sub

Private Sub imgNight_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgNight.Appearance = 1
    imgNight.BorderStyle = 1
End Sub

Private Sub imgNight_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgNight.Appearance = 0
    imgNight.BorderStyle = 0
End Sub

Private Sub imgOnOff_Click()
    If frmSysTray.STATE_SCREEN = True Then
        Unload frmBlackScreen
        imgOnOff.Picture = imgOff.Picture
    Else
        frmBlackScreen.Show
        frmBlackScreen.ZOrder 0
        Me.ZOrder 0
        imgOnOff.Picture = imgOn.Picture
    End If
End Sub

Private Sub imgSettingsOff_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgSettingsOn.Visible = True
    imgSettingsOff.Visible = False
End Sub

Private Sub imgSettingsOff_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Unload Me
    frmConfig.Show 0, frmBlackScreen
    If frmConfig.Visible = True Then
        imgSettingsOff.Visible = True
        imgSettingsOn.Visible = False
    End If
End Sub

Private Sub imgSettingsOn_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgSettingsOff.Visible = True
    imgSettingsOn.Visible = False
End Sub

Private Sub Timer1_Timer()
    If frmSysTray.STATE_SCREEN = True Then
        imgOnOff.Picture = imgOn.Picture
    Else
        imgOnOff.Picture = imgOff.Picture
    End If
    
    If ActiveApp = 0 Then
        ActiveApp = GetActiveWindow
        Timer2.Enabled = False
    End If
    
    If GetActiveWindow <> ActiveApp Then
        Timer2.Enabled = True
    End If
End Sub

Private Sub Timer2_Timer()
    If GetActiveWindow <> ActiveApp Then
        Unload Me
    End If
End Sub
