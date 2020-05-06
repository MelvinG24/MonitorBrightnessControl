VERSION 5.00
Begin VB.Form frmSysTray 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Form1"
   ClientHeight    =   675
   ClientLeft      =   1425
   ClientTop       =   2295
   ClientWidth     =   1680
   Icon            =   "frmSysTray.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   45
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   112
   ShowInTaskbar   =   0   'False
   Begin VB.Menu mPopupMenuMain 
      Caption         =   "SysTray"
      Visible         =   0   'False
      Begin VB.Menu mPopupMenu 
         Caption         =   "About"
         Index           =   0
      End
      Begin VB.Menu mPopupMenu 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mPopupMenu 
         Caption         =   "On/Off"
         Checked         =   -1  'True
         Index           =   2
      End
      Begin VB.Menu mPopupMenu 
         Caption         =   "Language"
         Enabled         =   0   'False
         Index           =   3
         Visible         =   0   'False
         Begin VB.Menu mLanguage 
            Caption         =   "Español (Es)"
            Index           =   0
         End
         Begin VB.Menu mLanguage 
            Caption         =   "English (En-US)"
            Checked         =   -1  'True
            Index           =   1
         End
      End
      Begin VB.Menu mPopupMenu 
         Caption         =   "Config"
         Index           =   4
      End
      Begin VB.Menu mPopupMenu 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mPopupMenu 
         Caption         =   "Exit"
         Index           =   6
      End
   End
End
Attribute VB_Name = "frmSysTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private WithEvents SysTray As clsSysTray
Attribute SysTray.VB_VarHelpID = -1
Public M_BRIGHTNESS As Integer
Public STATE_SCREEN As Boolean
Public shortCUT As Boolean
Public rBrightness As String
Public lBrightness As String

Private Sub Form_Load()
    rBrightness = "Ctrl + Shift + F5"
    lBrightness = "Ctrl + Shift + F6"
    shortCUT = True
    M_BRIGHTNESS = 128
    STATE_SCREEN = True
    'STATE_SCREEN = GetSetting("vbMonitorBrightnessControl", "Settings", "ScreenLabel", 0)
    
    Set SysTray = New clsSysTray
    Me.WindowState = 1
    DoEvents
    Me.Hide
    SysTray.Init Me, "Monitor Brightness Control"
    If STATE_SCREEN Then frmBlackScreen.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim frm As Form
    For Each frm In Forms
        Unload frm
        Set frm = Nothing
    Next frm
End Sub

Private Sub form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SysTray.MouseMove Button, X, Me
End Sub

Private Sub mPopupMenu_Click(Index As Integer)
    Select Case Me.mPopupMenu(Index).Caption
        Case "About":
            If frmControl.Visible = True Then
                Unload frmControl
            End If
                If frmBlackScreen.Visible = True Then
                    frmAbout.Show 0, frmBlackScreen
                Else
                    frmAbout.Show
                End If
        Case "On/Off":
            If STATE_SCREEN = True Then
                Me.mPopupMenu(Index).Checked = False
                STATE_SCREEN = False
                Unload frmBlackScreen
            ElseIf STATE_SCREEN = False Then
                Me.mPopupMenu(Index).Checked = True
                STATE_SCREEN = True
                frmBlackScreen.Show
                frmBlackScreen.ZOrder 0
            End If
        Case "Config":
            If frmBlackScreen.Visible = True Then
                frmConfig.Show 0, frmBlackScreen
            Else
                frmConfig.Show
            End If
        Case "Exit": Unload Me
        Case Else: MsgBox Me.mPopupMenu(Index).Caption
    End Select
End Sub

Private Sub SysTray_LeftClick()
    If frmAbout.Visible = True Then
        Unload frmAbout
    End If
    If frmConfig.Visible = False Then
        If frmControl.Visible = False Then
            frmControl.Show
            
            'Dim pt As POINTAPI
            'GetCursorPos pt
            'frmControl.Show
            'frmControl.Move pt.X * Screen.TwipsPerPixelX, (pt.Y * Screen.TwipsPerPixelY) - frmControl.Height
        End If
    Else
        Beep
        frmConfig.SetFocus
        MsgBox "You need first close the" & vbNewLine & "Monitor Brightness Control configuration window", vbInformation, "Info"
    End If
End Sub

Private Sub SysTray_RightClick()
    If STATE_SCREEN Then
        Me.mPopupMenu(2).Checked = True
    Else
        Me.mPopupMenu(2).Checked = False
    End If
    PopupMenu Me.mPopupMenuMain
End Sub
