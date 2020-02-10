VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmControl 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4185
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   975
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   975
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   0
      Top             =   3240
   End
   Begin MSComctlLib.Slider sliderControl 
      Height          =   3255
      Left            =   200
      TabIndex        =   2
      Top             =   120
      Width           =   600
      _ExtentX        =   1058
      _ExtentY        =   5741
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
      ScaleWidth      =   975
      TabIndex        =   0
      Top             =   3690
      Width           =   975
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
         ToolTipText     =   "Add custome shortcut"
         Top             =   120
         Width           =   630
      End
   End
End
Attribute VB_Name = "frmControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ActiveApp As Long
Dim r As Integer

Private Declare Function GetActiveWindow Lib "user32" () As Long

Private Sub Form_Load()
    r = (frmSysTray.M_BRIGHTNESS * 10) / 25.5
    sliderControl.Value = r
    Timer1.Enabled = True
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If X >= Me.Left And X <= Me.Left + Me.Width And _
'        Y >= Me.Top And Y <= Me.Top + Me.Height Then
'        MsgBox "Shape1 has been clicked."
'    End If
End Sub

Private Sub Form_LostFocus()
    Unload Me
End Sub

Private Sub lblConfig_Click()
    frmConfig.Show
    Unload Me
End Sub

Private Sub sliderControl_Change()
    r = (sliderControl.Value * 25.5) / 10
    frmSysTray.M_BRIGHTNESS = r
End Sub

Private Sub Timer1_Timer()
    If ActiveApp = 0 Then
        ActiveApp = GetActiveWindow
    End If
    
    If GetActiveWindow <> ActiveApp Then
        Timer1.Enabled = False
        ActiveApp = 0
        Unload Me
    End If
End Sub
