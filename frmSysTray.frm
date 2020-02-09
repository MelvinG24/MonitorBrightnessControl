VERSION 5.00
Begin VB.Form frmSysTray 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Form1"
   ClientHeight    =   675
   ClientLeft      =   1425
   ClientTop       =   2295
   ClientWidth     =   1680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   45
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   112
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmr 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   240
      Top             =   120
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   840
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   0
      Top             =   120
      Width           =   480
   End
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
         Caption         =   "Language"
         Index           =   2
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
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mPopupMenu 
         Caption         =   "Exit"
         Index           =   4
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
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set SysTray = Nothing
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SysTray.MouseMove Button, X, Me
End Sub

Private Sub pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SysTray.MouseMove Button, X, Me
End Sub

Private Sub mPopupMenu_Click(Index As Integer)
    Select Case Me.mPopupMenu(Index).Caption
        Case "About": frmAbout.Show vbModal
        Case "Exit": Unload Me
        Case Else: MsgBox Me.mPopupMenu(Index).Caption
    End Select
End Sub

Private Sub SysTray_BalloonClicked()
    MsgBox "Hola mundo", vbInformation, "Notice"
End Sub

Private Sub SysTray_LeftClick()
    frmControl.Show
    'SysTray.ShowBalloonTip "Right click for options", beInformation, "Tip"
End Sub

Private Sub SysTray_RightClick()
    PopupMenu Me.mPopupMenuMain
End Sub
