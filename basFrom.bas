Attribute VB_Name = "basFrom"
Option Explicit

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public WindowRect As RECT

Private Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" ( _
                ByVal hWnd As Long, _
                ByVal lpOperation As String, _
                ByVal lpFile As String, _
                ByVal lpParameters As String, _
                ByVal lpDirectory As String, _
                ByVal nShowCmd As Long) As Long
                
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" ( _
                ByVal uAction As Long, _
                ByVal uParam As Long, _
                ByRef lpvParam As Any, _
                ByVal fuWinIni As Long) As Long
                
Public Const SPI_GETWORKAREA = 48
Public Const rB_SC As String = "Ctrl + Shift + F5"
Public Const lB_SC As String = "Ctrl + Shift + F6"
Public M_BRIGHTNESS As Integer
Public SHOW_SHORTCUTS As Boolean
Public rBrightness As String
Public lBrightness As String

Private Sub Main()
    If App.PrevInstance Then Exit Sub
    
    rBrightness = rB_SC
    lBrightness = lB_SC
    M_BRIGHTNESS = 128
    SHOW_SHORTCUTS = True
    
    'Get monitor work area size -without taskbar or desktop toolbars obstruction
    SystemParametersInfo SPI_GETWORKAREA, 0, WindowRect, 0
    
    Load frmSysTray
End Sub

Public Sub HuperJump(ByVal URL As String)
    Call ShellExecute(0&, vbNullString, URL, vbNullString, vbNullString, vbNormalFocus)
End Sub

Public Function showAbout()
    If frmControl.Visible Then Unload frmControl
    If frmMain.Visible Then
        frmAbout.Show 1, frmMain
    Else
        frmAbout.Show
    End If
End Function

Public Function OnOffSwitch(Index As Integer)
    If frmSysTray.mPopupMenu(Index).Checked = True Then
        frmSysTray.mPopupMenu(Index).Checked = False
        Unload frmMain
    Else
        frmSysTray.mPopupMenu(Index).Checked = True
        frmMain.Show
    End If
End Function

Public Function showControl()
    If frmAbout.Visible Then Unload frmAbout
    If frmConfig.Visible = False Then
        If frmControl.Visible = False Then
            If frmMain.Visible Then
                frmControl.Show 1, frmMain
            Else
                frmControl.Show
            End If
        End If
    Else
        Beep
        frmConfig.SetFocus
        MsgBox "You need first close the" & vbNewLine & "Monitor Brightness Control configuration window", vbInformation, "Info"
    End If
End Function

Public Function showConfig()
    If frmMain.Visible Then
        frmConfig.Show 1, frmMain
    Else
        frmConfig.Show
    End If
End Function

Public Sub timerOnOff(b As Boolean)
    If frmMain.Visible Then frmMain.Timer2.Enabled = b
End Sub
