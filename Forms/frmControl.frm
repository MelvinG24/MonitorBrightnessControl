VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MsComCtl.ocx"
Begin VB.Form frmControl 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4320
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   1020
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   4320
   ScaleMode       =   0  'User
   ScaleWidth      =   1098.462
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   0
      Top             =   3240
   End
   Begin MSComctlLib.Slider sliderControl 
      Height          =   2295
      Left            =   175
      TabIndex        =   2
      Top             =   720
      Width           =   600
      _ExtentX        =   1058
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
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   1020
      TabIndex        =   0
      Top             =   3825
      Width           =   1020
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
         MouseIcon       =   "frmControl.frx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   1
         ToolTipText     =   "App Configurations"
         Top             =   120
         Width           =   630
      End
   End
   Begin VB.Image imgDay 
      Height          =   500
      Left            =   240
      MouseIcon       =   "frmControl.frx":030A
      MousePointer    =   99  'Custom
      Picture         =   "frmControl.frx":0614
      Stretch         =   -1  'True
      Top             =   120
      Width           =   500
   End
   Begin VB.Image imgNight 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   240
      MouseIcon       =   "frmControl.frx":52E01
      MousePointer    =   99  'Custom
      Picture         =   "frmControl.frx":5310B
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

Dim sys As StatusBar
Dim ActiveApp As Long
Dim r As Integer

Private Declare Function GetActiveWindow Lib "user32" () As Long

Private Sub Form_Activate()
    If Me.Visible Then
        'Load Toll-Tip string from RES
        lblConfig.TooltipText = LoadResString(109)
        
        ActiveApp = 0
        r = (P_VarBrightnessLevel * 10) / 25.5
        sliderControl.Value = r
        Timer1.Enabled = True
        SystemParametersInfo SPI_GETWORKAREA, 0, WindowRect, 0
        frmMain.Timer1.Enabled = True
    End If
End Sub

Private Sub Form_GotFocus()
    ActiveApp = GetActiveWindow
End Sub

Private Sub Form_LostFocus()
    Unload Me
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X >= Me.Left And X <= Me.Left + Me.Width And _
        Y >= Me.Top And Y <= Me.Top + Me.Height Then
        MsgBox "Shape1 has been clicked."
    End If
End Sub

Private Sub Form_Resize()
    Me.Top = (WindowRect.Bottom * Screen.TwipsPerPixelY - Me.Height) - 120
    Me.Left = (WindowRect.Right * Screen.TwipsPerPixelX - Me.Width) - 120
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Timer1.Enabled = False
    frmMain.Timer1.Enabled = False
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

Private Sub lblConfig_Click()
    Unload Me
    showConfig
End Sub

Private Sub sliderControl_Change()
    r = (sliderControl.Value * 25.5) / 10
    P_VarBrightnessLevel = r
End Sub

Private Sub sliderControl_Scroll()
    sliderControl_Change
End Sub

Private Sub Timer1_Timer()
    If ActiveApp = 0 Then
        ActiveApp = GetActiveWindow
    End If

    If GetActiveWindow <> ActiveApp Then
        ActiveApp = 0
        Unload Me
    End If
End Sub
