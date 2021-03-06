VERSION 5.00
Begin VB.Form frmSysTray 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Form1"
   ClientHeight    =   675
   ClientLeft      =   120
   ClientTop       =   450
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
         Index           =   2
      End
      Begin VB.Menu mPopupMenu 
         Caption         =   "Brightness"
         Index           =   3
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

Private WithEvents SysTray As clsSysTray
Attribute SysTray.VB_VarHelpID = -1

'----------------------------------------------------------
' General Functions
'----------------------------------------------------------
'Mouse left-click on project icon
Private Sub SysTray_LeftClick()
    showControl
End Sub

'Mouse right-click on project icon
Private Sub SysTray_RightClick()
    If frmAbout.Visible Then Unload frmAbout
    If frmControl.Visible Then Unload frmControl
    If frmConfig.Visible Then
        frmConfig.SetFocus
    Else
        PopupMenu Me.mPopupMenuMain
    End If
End Sub

'Specify popup-menu language
Public Sub chgPopLng()
    Dim I, M As Integer
    M = 0
    For I = 0 To Me.mPopupMenu.UBound
        If Not Me.mPopupMenu(I).Caption = "-" Then
            Me.mPopupMenu(I).Caption = LoadResString(103 + M + L)
            M = M + 1
        End If
    Next I
End Sub

'----------------------------------------------------------
' Form/Controls Actions
'----------------------------------------------------------
Private Sub Form_Load()
'    Me.Icon = LoadResData(101 + P_ICON, "CUSTOM")
'    If P_ICON = 1 Then Me.Icon = imgWin7.Picture       'Element imgWin7 is now deleted
    chgPopLng
    Set SysTray = New clsSysTray
    Me.WindowState = 1
    DoEvents
    Me.Hide
    SysTray.Init Me, "Monitor Brightness Control"
    
    If P_VarChckRunBS = 1 Then
        Me.mPopupMenu(2).Checked = True
        frmMain.Show
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SysTray.MouseMove Button, X, Me
End Sub

Private Sub mPopupMenu_Click(Index As Integer)
    Select Case Me.mPopupMenu(Index).Index
        Case 0: showAbout
        Case 2: OnOffSwitch Index
        Case 3: showControl
        Case 4: showConfig
        Case 6: Unload Me
        Case Else: MsgBox Me.mPopupMenu(Index).Caption
    End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim frm As Form
    For Each frm In Forms
        Unload frm
        Set frm = Nothing
    Next frm
End Sub
