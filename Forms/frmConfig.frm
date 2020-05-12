VERSION 5.00
Begin VB.Form frmConfig 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configurations"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5550
   Icon            =   "frmConfig.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   5550
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.ComboBox cmdLanguage 
      Height          =   315
      ItemData        =   "frmConfig.frx":030A
      Left            =   960
      List            =   "frmConfig.frx":0314
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   4320
      Width           =   1695
   End
   Begin VB.CheckBox chEnableSC 
      Caption         =   "Enable"
      Height          =   495
      Left            =   4440
      TabIndex        =   1
      Top             =   0
      Value           =   1  'Checked
      Width           =   855
   End
   Begin VB.CheckBox chLabel 
      Caption         =   "On/Off on-screen shortcuts label"
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   3720
      Value           =   1  'Checked
      Width           =   4815
   End
   Begin VB.CommandButton btnDone 
      Caption         =   "Done"
      Height          =   495
      Left            =   4080
      TabIndex        =   12
      ToolTipText     =   "Save all changes"
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   120
      TabIndex        =   15
      Top             =   2640
      Width           =   5295
      Begin VB.CheckBox chOnOff 
         Caption         =   "When program started run Black-Screen on mode = ON"
         Enabled         =   0   'False
         Height          =   255
         Left            =   720
         TabIndex        =   9
         Top             =   600
         Width           =   4335
      End
      Begin VB.CheckBox chStartUp 
         Caption         =   "Start-Up program with MS Windows"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   4815
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
         Index           =   0
         X1              =   360
         X2              =   360
         Y1              =   480
         Y2              =   720
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Brightness Shortcuts"
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5295
      Begin VB.CommandButton btnSettings 
         Cancel          =   -1  'True
         Caption         =   "Change"
         Height          =   375
         Index           =   1
         Left            =   1800
         TabIndex        =   4
         ToolTipText     =   "Activate shortcut TextBox"
         Top             =   1065
         Width           =   765
      End
      Begin VB.CommandButton btnSettings 
         Caption         =   "Save"
         Height          =   375
         Index           =   3
         Left            =   3960
         TabIndex        =   7
         ToolTipText     =   "Save shortcuts settings"
         Top             =   1800
         Width           =   1000
      End
      Begin VB.CommandButton btnSettings 
         Caption         =   "Change"
         Height          =   375
         Index           =   0
         Left            =   1800
         TabIndex        =   2
         ToolTipText     =   "Activate shortcut TextBox"
         Top             =   450
         Width           =   765
      End
      Begin VB.CommandButton btnSettings 
         Caption         =   "Reset"
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   6
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
         Left            =   2520
         TabIndex        =   5
         Top             =   1065
         Width           =   2475
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
         Left            =   2520
         TabIndex        =   3
         Top             =   450
         Width           =   2475
      End
      Begin VB.Shape Shape1 
         FillStyle       =   5  'Downward Diagonal
         Height          =   2415
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   5295
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000A&
         X1              =   240
         X2              =   5040
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   240
         TabIndex        =   14
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
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   240
         TabIndex        =   13
         Top             =   500
         Width           =   840
      End
   End
   Begin VB.Label lblLanguage 
      AutoSize        =   -1  'True
      Caption         =   "Language:"
      Height          =   195
      Left            =   120
      TabIndex        =   16
      Top             =   4350
      Width           =   780
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000A&
      Index           =   1
      X1              =   120
      X2              =   5280
      Y1              =   4080
      Y2              =   4080
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents SysTray As clsSysTray
Attribute SysTray.VB_VarHelpID = -1

Private Sub btnSettings_Click(Index As Integer)
    Select Case btnSettings(Index).Index
        Case 0:
            txtBrightUp.Enabled = True
            txtBrightUp.Text = ""
            txtBrightUp.SetFocus
        Case 1:
            txtBrightDown.Enabled = True
            txtBrightDown.Text = ""
            txtBrightDown.SetFocus
        Case 2:
            txtBrightUp.Text = rB_SC
            txtBrightUp.Enabled = False
            txtBrightDown.Text = lB_SC
            txtBrightDown.Enabled = False
            
            rBrightness = txtBrightUp.Text
            lBrightness = txtBrightDown.Text
        Case 3:
default:
        MsgBox "Button do not exists"
    End Select
End Sub

Private Sub btnDone_Click()
    'Check if the shortcut match with the preview shortcut
    'If they do not match and it still not saved, show a msgbox asking for saving
    Unload Me
End Sub

Private Sub chEnableSC_Click()
    Dim Check, Index As Integer
    
    Check = chEnableSC.Value
    
    Frame1.Enabled = Check
    Shape1.Visible = CBool(Check = 0)
    For Index = 0 To btnSettings.UBound
        btnSettings(Index).Enabled = Check
    Next
End Sub

Private Sub chLabel_Click()
    Dim Check As Integer
    
    Check = chLabel.Value
    
    frmMain.lblInfo.Visible = Check
    SHOW_SHORTCUTS = Check
End Sub

Private Sub chStartUp_Click()
    Dim Check As Integer
    
    Check = chStartUp.Value
    
    chStartUp.FontBold = Check
    chOnOff.Enabled = Check
    
    If chStartUp.Value = 0 Then chOnOff.Value = 0
    
    If m_IgnoreEvents Then Exit Sub
    SetRunAtStartUp App.EXEName, App.Path
End Sub

Private Sub cmdLanguage_Click()
    chckLanguage = cmdLanguage.ListIndex
End Sub

Private Sub Form_Activate()
    txtBrightUp.Text = rBrightness
    txtBrightDown.Text = lBrightness
    chStartUp.Value = chckRunAtStartUp
    chOnOff.Value = chckRunAfter
    cmdLanguage.ListIndex = chckLanguage
'    chLabel.Value = SHOW_SHORTCUTS
    timerOnOff False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_LostFocus()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    timerOnOff True
End Sub

Private Sub txtBrightUp_KeyDown(KeyCode As Integer, Shift As Integer)
'    txtBrightUp.Text
End Sub

Private Sub txtBrightUp_LostFocus()
    If txtBrightUp.Text = "" Then
        txtBrightUp.Text = rBrightness
        txtBrightUp.Enabled = False
    End If
End Sub
