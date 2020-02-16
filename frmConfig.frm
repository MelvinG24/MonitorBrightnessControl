VERSION 5.00
Begin VB.Form frmConfig 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Configurations"
   ClientHeight    =   4560
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6570
   Icon            =   "frmConfig.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   6570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.PictureBox shHoldFrm 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   100
      MouseIcon       =   "frmConfig.frx":7F6A
      MousePointer    =   99  'Custom
      ScaleHeight     =   345
      ScaleWidth      =   705
      TabIndex        =   14
      ToolTipText     =   "Hold to move the form"
      Top             =   3960
      Width           =   705
      Begin VB.Shape shLines 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808080&
         FillStyle       =   2  'Horizontal Line
         Height          =   345
         Left            =   0
         Top             =   0
         Width           =   705
      End
   End
   Begin VB.CheckBox chLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "On/Off on-screen shortcuts label"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1320
      TabIndex        =   9
      Top             =   3960
      Value           =   1  'Checked
      Width           =   2775
   End
   Begin VB.CommandButton btnDone 
      Caption         =   "Done"
      Height          =   495
      Left            =   5160
      TabIndex        =   10
      ToolTipText     =   "Save all changes"
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   1080
      TabIndex        =   13
      Top             =   2640
      Width           =   5295
      Begin VB.CheckBox chOnOff 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "By default start black screen ON"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   720
         TabIndex        =   8
         Top             =   600
         Width           =   2655
      End
      Begin VB.CheckBox chStartUp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Start-Up automatly with MS Windows"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   4695
      End
      Begin VB.Line Line4 
         BorderStyle     =   3  'Dot
         X1              =   600
         X2              =   360
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line Line3 
         BorderStyle     =   3  'Dot
         X1              =   360
         X2              =   360
         Y1              =   480
         Y2              =   720
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Brightness Shortcuts"
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   5295
      Begin VB.CommandButton btnChangeBrightDown 
         Cancel          =   -1  'True
         Caption         =   "Change"
         Height          =   375
         Left            =   2230
         TabIndex        =   3
         ToolTipText     =   "Activate shortcut TextBox"
         Top             =   1065
         Width           =   765
      End
      Begin VB.CommandButton btnSave 
         Caption         =   "Save"
         Height          =   375
         Left            =   3960
         TabIndex        =   6
         ToolTipText     =   "Save shortcuts settings"
         Top             =   1800
         Width           =   1000
      End
      Begin VB.CommandButton btnChangeBrightUp 
         Caption         =   "Change"
         Height          =   375
         Left            =   2230
         TabIndex        =   1
         ToolTipText     =   "Activate shortcut TextBox"
         Top             =   450
         Width           =   765
      End
      Begin VB.CommandButton btnReset 
         Caption         =   "Reset"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         ToolTipText     =   "Reset shortcuts setting to their defaults values"
         Top             =   1800
         Width           =   1000
      End
      Begin VB.TextBox txtBrightDown 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Ubuntu Mono"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3000
         TabIndex        =   4
         Text            =   "Alt F6"
         Top             =   1065
         Width           =   2000
      End
      Begin VB.TextBox txtBrightUp 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Ubuntu Mono"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3000
         TabIndex        =   2
         Text            =   "Shift F5"
         Top             =   450
         Width           =   2000
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000A&
         X1              =   240
         X2              =   5040
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Bright-Down"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   240
         TabIndex        =   12
         Top             =   1120
         Width           =   1080
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000A&
         X1              =   240
         X2              =   5040
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Bright-Up"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   240
         TabIndex        =   11
         Top             =   500
         Width           =   840
      End
   End
   Begin VB.Shape shActiveOp 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   500
      Left            =   800
      Top             =   200
      Width           =   100
   End
   Begin VB.Image imgGear 
      Appearance      =   0  'Flat
      Height          =   400
      Left            =   250
      Picture         =   "frmConfig.frx":80BC
      Stretch         =   -1  'True
      Top             =   250
      Width           =   400
   End
   Begin VB.Shape shBorder 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   4560
      Left            =   0
      Top             =   0
      Width           =   900
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lScreen As Boolean

Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long

Private WithEvents SysTray As clsSysTray
Attribute SysTray.VB_VarHelpID = -1

Private Sub btnChangeBrightDown_Click()
    txtBrightDown.Enabled = True
    txtBrightDown.Text = ""
    txtBrightDown.SetFocus
End Sub

Private Sub btnChangeBrightUp_Click()
    txtBrightUp.Enabled = True
    txtBrightUp.Text = ""
    txtBrightUp.SetFocus
End Sub

Private Sub btnDone_Click()
    'Check if the shortcut match with the preview shortcut
    'If they do not match and it still not saved, show a msgbox asking for saving
    'SaveSetting "vbMonitorBrightnessControl", "Settings", "ScreenLabel", lScreen
    'SaveSetting "CounterApp", "Counts", "LastTime", Now
    Unload Me
End Sub

Private Sub btnReset_Click()
    txtBrightUp.Text = "Ctrl + Shift + F5"
    txtBrightUp.Enabled = False
    txtBrightDown.Text = "Ctrl + Alt + F6"
    txtBrightDown.Enabled = False
End Sub

Private Sub chLabel_Click()
    If chLabel.Value = 1 Then
        lScreen = True
    ElseIf chLabel.Value = 0 Then
        lScreen = False
    End If
End Sub

Private Sub chStartUp_Click()
    If chStartUp.Value = 1 Then
        chStartUp.FontBold = True
        chOnOff.Enabled = True
    ElseIf chStartUp.Value = 0 Then
        chStartUp.FontBold = False
        chOnOff.Enabled = False
        chOnOff.Value = 0
    End If
End Sub

Private Sub Form_Load()
    frmRounded Me
    lScreen = frmSysTray.STATE_SCREEN
    If Me.Visible = True Then
        frmMain.Timer2.Enabled = False
        If lScreen Then
            chLabel.Value = 1
        Else
            chLabel.Value = 0
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.Timer2.Enabled = True
End Sub

Private Sub shHoldFrm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Formulario Me
End Sub

Private Sub txtBrightUp_GotFocus()
    'SysTray.ShowBalloonTip "Press your shortcut combination now", beInformation, "Tip"
End Sub

Public Sub Formulario(frm As Form)
    ReleaseCapture
    Call SendMessage(frm.hwnd, &HA1, 2, 0&)
End Sub

