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
      ItemData        =   "frmConfig.frx":1CFA
      Left            =   960
      List            =   "frmConfig.frx":1D04
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   4320
      Width           =   1695
   End
   Begin VB.CheckBox ChckSCEnable 
      Caption         =   "Enable"
      Height          =   495
      Left            =   4320
      TabIndex        =   1
      Top             =   0
      Value           =   1  'Checked
      Width           =   975
   End
   Begin VB.CheckBox ChckSCVisible 
      Caption         =   "On/Off on-screen shortcuts label"
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   3720
      Value           =   1  'Checked
      Width           =   4815
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   4200
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
      Begin VB.CheckBox ChckRunBS 
         Caption         =   "When program started run Black-Screen on mode = ON"
         Enabled         =   0   'False
         Height          =   255
         Left            =   720
         TabIndex        =   9
         Top             =   600
         Width           =   4335
      End
      Begin VB.CheckBox ChckRunStartUp 
         Caption         =   "Start-Up program with MS Windows"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   4815
      End
      Begin VB.Line Line1 
         BorderStyle     =   3  'Dot
         Index           =   1
         X1              =   600
         X2              =   360
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line Line1 
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
         Left            =   1750
         TabIndex        =   4
         Top             =   1065
         Width           =   765
      End
      Begin VB.CommandButton btnSettings 
         Caption         =   "Set"
         Height          =   375
         Index           =   3
         Left            =   3960
         TabIndex        =   7
         ToolTipText     =   "Set shortcuts change"
         Top             =   1800
         Width           =   1000
      End
      Begin VB.CommandButton btnSettings 
         Caption         =   "Change"
         Height          =   375
         Index           =   0
         Left            =   1750
         TabIndex        =   2
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
            Name            =   "Arial"
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
            Name            =   "Arial"
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
      Begin VB.Line Line 
         BorderColor     =   &H00808080&
         Index           =   1
         X1              =   240
         X2              =   5040
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Label lblBrightDown 
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
      Begin VB.Line Line 
         BorderColor     =   &H00808080&
         Index           =   0
         X1              =   240
         X2              =   5040
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label lblBrightUp 
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
   Begin VB.Line Line 
      BorderColor     =   &H00808080&
      Index           =   2
      X1              =   120
      X2              =   5400
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

'% Integer
'& Long
'! Single
'# Double
'$ String
'@ Currency

Private WithEvents SysTray As clsSysTray
Attribute SysTray.VB_VarHelpID = -1

'----------------------------------------------------------
' General Functions
'----------------------------------------------------------
'Change language
Private Sub chgLanguage()
    Me.Caption = LoadResString(108 + L)
    'Brightness Shortcuts frame
    Frame1.Caption = LoadResString(109 + L)
    ChckSCEnable.Caption = LoadResString(110 + L)
    ChckSCEnable.TooltipText = LoadResString(111 + L)
    lblBrightUp.Caption = LoadResString(112 + L)
    lblBrightDown.Caption = LoadResString(113 + L)
    btnSettings(0).Caption = LoadResString(114 + L)
    btnSettings(0).TooltipText = LoadResString(115 + L)
    btnSettings(1).Caption = LoadResString(114 + L)
    btnSettings(1).TooltipText = LoadResString(115 + L)
    btnSettings(2).Caption = LoadResString(116 + L)
    btnSettings(2).TooltipText = LoadResString(117 + L)
    btnSettings(3).Caption = LoadResString(118 + L)
    btnSettings(3).TooltipText = LoadResString(119 + L)
    'Startup program frame
    ChckRunStartUp.Caption = LoadResString(122 + L)
    ChckRunBS.Caption = LoadResString(123 + L)
    'Others settings options
    ChckSCVisible.Caption = LoadResString(124 + L)
    lblLanguage.Caption = LoadResString(125 + L)
    cmdLanguage.TooltipText = LoadResString(126 + L)
    btnSave.Caption = LoadResString(127 + L)
    btnSave.TooltipText = LoadResString(128 + L)
    'Call systray and black-screen forms to change language
    frmSysTray.chgPopLng
    If frmMain.Visible Then frmMain.chgMainLblLng
End Sub

'If Short-cut textbox has lost focus
Private Function F_txtBright_lostFocus(txt As TextBox) As Boolean
    F_txtBright_lostFocus = False
    If txt.Text = "" Then
        If txt.Name = txtBrightUp.Name Then
            txt.Text = P_VarRsBrightness
        Else
            txt.Text = P_VarLwBrightness
        End If
        txt.Enabled = False
    Else
        If txt.Enabled Then
            MsgBox LoadResString(121 + L), vbInformation, LoadResString(120 + L)
            btnSettings(3).SetFocus
            F_txtBright_lostFocus = True
        End If
    End If
End Function

'If Short-cut btn is clicked
Private Sub F_chckSCBtn(txt As TextBox, lostF As TextBox)
    If Not F_txtBright_lostFocus(lostF) Then 'Check if any of the SC_TextBox has lost focus
        With txt
            .Enabled = True
            .Text = ""
            .SetFocus
        End With
    End If
End Sub

'If Set btn is clicked
Private Function F_setSCBtn(txt As TextBox) As Boolean
    F_setSCBtn = False
    If txt.Text <> "" Then
        txt.Enabled = False
    Else
        F_setSCBtn = True
        MsgBox LoadResString(129 + L)
        txt.SetFocus
    End If
End Function

'----------------------------------------------------------
' Form/Controls Actions
'----------------------------------------------------------
Private Sub btnSettings_Click(Index As Integer)
    Select Case btnSettings(Index).Index
        Case 0: 'Btn bright-up short-cut change
            F_chckSCBtn txtBrightUp, txtBrightDown
            Exit Sub
        Case 1: 'Btn bright-down short-cut change
            F_chckSCBtn txtBrightDown, txtBrightUp
            Exit Sub
        Case 2: 'Reset btn
            txtBrightUp.Text = rB_SC
            txtBrightUp.Enabled = False
            txtBrightDown.Text = lB_SC
            txtBrightDown.Enabled = False
            Exit Sub
        Case 3: 'Set btn
            If txtBrightUp.Enabled Then F_setSCBtn txtBrightUp
            If txtBrightDown.Enabled Then F_setSCBtn txtBrightDown
    End Select
End Sub

Private Sub btnSave_Click()
    'Check if the shortcut match with the preview shortcut
    'If they do not match and it still not saved, show a msgbox asking for saving
    Unload Me
End Sub

Private Sub ChckSCEnable_Click()
'*********************************
'  Need to fix the double MsgBox
'*********************************
    Dim Check, Index As Integer
    
    Check = ChckSCEnable.Value
    
    If F_setSCBtn(txtBrightUp) Or F_setSCBtn(txtBrightDown) Then
        If ChckSCEnable.Value = 0 Then ChckSCEnable.Value = 1
        Exit Sub
    Else
        Frame1.Enabled = Check
        Shape1.Visible = CBool(Check = 0)
        For Index = 0 To btnSettings.UBound
            btnSettings(Index).Enabled = Check
        Next
    End If
End Sub

Private Sub ChckSCVisible_Click()
    Dim Check As Integer
    
    Check = ChckSCVisible.Value
    
    frmMain.lblInfo.Visible = Check
End Sub

Private Sub ChckRunStartUp_Click()
    Dim Check As Integer
    
    Check = ChckRunStartUp.Value
    
    ChckRunStartUp.FontBold = Check
    ChckRunBS.Enabled = Check
  
    If m_IgnoreEvents Then Exit Sub
    SetRunAtStartUp App.EXEName, App.Path, Check
End Sub

Private Sub cmdLanguage_Click()
'    P_VarChckLanguage = cmdLanguage.ListIndex
    F_L cmdLanguage.ListIndex
    chgLanguage
End Sub

Private Sub Form_Activate()
    If Me.Visible Then
        timerOnOff False
'        Me.Icon = LoadResPicture(101 + P_ICON, vbResIcon)
        txtBrightDown.Text = P_VarLwBrightness
        txtBrightUp.Text = P_VarRsBrightness
        ChckSCEnable.Value = P_VarChckSCEnable
        ChckRunStartUp.Value = P_VarChckRunStartUp
        ChckRunBS.Value = P_VarChckRunBS
        ChckSCVisible.Value = P_VarChckSCVisible
        cmdLanguage.ListIndex = P_VarChckLanguage 'This action by default change program language
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_LostFocus()
    Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Me.Visible Then
        If Not F_setSCBtn(txtBrightUp) And Not F_setSCBtn(txtBrightDown) Then
            SaveSettings
            timerOnOff True
        Else
            Cancel = True
        End If
    End If
End Sub
