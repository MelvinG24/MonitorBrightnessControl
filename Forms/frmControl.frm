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
      Max             =   90
      SelStart        =   45
      TickStyle       =   2
      TickFrequency   =   45
      Value           =   45
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
         Alignment       =   2  'Center
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
   Begin VB.Image btnDayNight 
      Height          =   500
      Index           =   0
      Left            =   240
      MouseIcon       =   "frmControl.frx":030A
      MousePointer    =   99  'Custom
      Picture         =   "frmControl.frx":0614
      Stretch         =   -1  'True
      Top             =   120
      Width           =   500
   End
   Begin VB.Image btnDayNight 
      Appearance      =   0  'Flat
      Height          =   495
      Index           =   1
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

Dim ActiveApp As Long

Private Declare Function GetActiveWindow Lib "user32" () As Long

'----------------------------------------------------------
' Form/Controls Actions
'----------------------------------------------------------
Private Sub btnDayNight_Click(Index As Integer)
    Select Case btnDayNight(Index).Index
        Case 0: 'btnDay
            If sliderControl.Value >= 69 And sliderControl.Value <= 90 Then
                sliderControl.Value = 68
            ElseIf sliderControl.Value >= 47 And sliderControl.Value <= 68 Then
                sliderControl.Value = 46
            ElseIf sliderControl.Value >= 25 And sliderControl.Value <= 46 Then
                sliderControl.Value = 24
            ElseIf sliderControl.Value <= 24 Then
                sliderControl.Value = 0
            End If
        Case 1: 'btnNight
            If sliderControl.Value >= 0 And sliderControl.Value <= 23 Then
                sliderControl.Value = 24
            ElseIf sliderControl.Value >= 24 And sliderControl.Value <= 45 Then
                sliderControl.Value = 46
            ElseIf sliderControl.Value >= 46 And sliderControl.Value <= 67 Then
                sliderControl.Value = 68
            ElseIf sliderControl.Value >= 68 Then
                sliderControl.Value = 90
            End If
    End Select
End Sub

Private Sub btnDayNight_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    btnDayNight(Index).Appearance = 1
    btnDayNight(Index).BorderStyle = 1
End Sub

Private Sub btnDayNight_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    btnDayNight(Index).Appearance = 0
    btnDayNight(Index).BorderStyle = 0
End Sub

Private Sub Form_Activate()
    If Me.Visible Then
        'Load Toll-Tip string from RES
        lblConfig.TooltipText = LoadResString(108 + L)
        
        ActiveApp = 0
        sliderControl.Value = (P_VarBrightnessLevel * 10) / 25.5
        Timer1.Enabled = True
        
        'Set focus on slider control
        sliderControl.SetFocus
    End If
End Sub

Private Sub Form_GotFocus()
    ActiveApp = GetActiveWindow
End Sub

Private Sub Form_LostFocus()
    Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Me.Visible Then
        Timer1.Enabled = False
    End If
End Sub

Private Sub Form_Resize()
    Me.Top = (WindowRect.Bottom * Screen.TwipsPerPixelY - Me.Height) - 120
    Me.Left = (WindowRect.Right * Screen.TwipsPerPixelX - Me.Width) - 120
End Sub

Private Sub lblConfig_Click()
    Unload Me
    showConfig
End Sub

Private Sub sliderControl_Change()
    P_VarBrightnessLevel = (sliderControl.Value * 25.5) / 10
    frmMain.SetWinTrans P_VarBrightnessLevel
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
