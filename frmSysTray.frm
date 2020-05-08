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
         Caption         =   "Brightness"
         Index           =   3
      End
      Begin VB.Menu mPopupMenu 
         Caption         =   "Language"
         Enabled         =   0   'False
         Index           =   4
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
         Index           =   5
      End
      Begin VB.Menu mPopupMenu 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mPopupMenu 
         Caption         =   "Exit"
         Index           =   7
      End
   End
End
Attribute VB_Name = "frmSysTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents SysTray As clsSysTray
Attribute SysTray.VB_VarHelpID = -1

Private Sub Form_Load()
    Set SysTray = New clsSysTray
    Me.WindowState = 1
    DoEvents
    Me.Hide
    SysTray.Init Me, "Monitor Brightness Control"
    
    If Me.mPopupMenu(2).Checked Then frmMain.Show
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SysTray.MouseMove Button, X, Me
End Sub

Private Sub mPopupMenu_Click(Index As Integer)
    Select Case Me.mPopupMenu(Index).Caption
        Case "About": showAbout
        Case "On/Off": OnOffSwitch Index
        Case "Brightness": showControl
        Case "Config": showConfig
        Case "Exit": unloadMe
        Case Else: MsgBox Me.mPopupMenu(Index).Caption
    End Select
End Sub

Private Sub SysTray_LeftClick()
    showControl
End Sub

Private Sub SysTray_RightClick()
    PopupMenu Me.mPopupMenuMain
End Sub

Private Sub unloadMe()
    Dim frm As Form
    For Each frm In Forms
        Unload frm
        Set frm = Nothing
    Next frm
End Sub
