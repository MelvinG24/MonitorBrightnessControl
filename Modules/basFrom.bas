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
                ByVal hwnd As Long, _
                ByVal lpOperation As String, _
                ByVal lpFile As String, _
                ByVal lpParameters As String, _
                ByVal lpDirectory As String, _
                ByVal nShowCmd As Long) As Long

Private Declare Function PathFileExists Lib "shlwapi" Alias "PathFileExistsA" ( _
                ByVal pszPath As String) As Long
                
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" ( _
                ByVal uAction As Long, _
                ByVal uParam As Long, _
                ByRef lpvParam As Any, _
                ByVal fuWinIni As Long) As Long
'
'Public Declare Function GetKeyState Lib "user32" ( _
'                ByVal nVirtKey As Long) As Integer

Public Const SPI_GETWORKAREA = 48
Public Const rB_SC As String = "Ctrl + Shift + F5"
Public Const lB_SC As String = "Ctrl + Shift + F6"
'Brightness variable
Public M_BRIGHTNESS As Integer
Public SHOW_SHORTCUTS As Boolean
Public rBrightness As String
Public lBrightness As String
'Check run-program variable
Public chckRunAtStartUp As Integer
Public chckRunAfter As Integer
'Language variable
Public chckLanguage As Integer

'----------------------------------------------------------
' Start-up program
'----------------------------------------------------------
Private Sub Main()
    If App.PrevInstance Then Exit Sub
    
    'Check if settings files exist
    If fileExistsCheck("Settings.dat") Then
        LoadSettings
    Else
        createSettings
    End If
    
'    rBrightness = rB_SC
'    lBrightness = lB_SC
'    M_BRIGHTNESS = 128
'    SHOW_SHORTCUTS = True
'    chckLanguage = 0
    
    m_IgnoreEvents = True
    If StartUp(App.EXEName) Then
        chckRunAtStartUp = 1
    Else
        chckRunAtStartUp = 0
        chckRunAfter = 0
    End If
    m_IgnoreEvents = False
    
    '¿Estara de Más aquí?
    'Get monitor work area size -without taskbar or desktop toolbars obstruction
    SystemParametersInfo SPI_GETWORKAREA, 0, WindowRect, 0
    
    Load frmSysTray
End Sub

'----------------------------------------------------------
' Hyperlink label
'----------------------------------------------------------
Public Sub HuperJump(ByVal URL As String)
    Call ShellExecute(0&, vbNullString, URL, vbNullString, vbNullString, vbNormalFocus)
End Sub

'----------------------------------------------------------
' Pop-up menu commands
'----------------------------------------------------------
Public Sub showAbout()
    If frmControl.Visible Then Unload frmControl
    If frmMain.Visible Then
        timerOnOff False
        frmAbout.Show 0, frmMain
    Else
        frmAbout.Show
    End If
End Sub

Public Sub OnOffSwitch(Index As Integer)
    If frmSysTray.mPopupMenu(Index).Checked = True Then
        frmSysTray.mPopupMenu(Index).Checked = False
        Unload frmMain
    Else
        frmSysTray.mPopupMenu(Index).Checked = True
        frmMain.Show
    End If
End Sub

Public Sub showControl()
    If frmAbout.Visible Then Unload frmAbout
    If frmConfig.Visible = False Then
        If frmControl.Visible = False Then
            If frmMain.Visible Then
                timerOnOff False
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
End Sub

Public Sub showConfig()
    If frmMain.Visible Then
        frmConfig.Show 0, frmMain
    Else
        frmConfig.Show
    End If
End Sub

'----------------------------------------------------------
' Main-form timer status
'----------------------------------------------------------
Public Sub timerOnOff(b As Boolean)
    If frmMain.Visible Then frmMain.Timer2.Enabled = b
End Sub

'----------------------------------------------------------
' Data-Settings functions
'----------------------------------------------------------
Private Sub createSettings()
    Dim f As Integer
    
    On Error Resume Next: Kill "Settings.dat": On Error GoTo 0
    f = FreeFile(0)
    Open "Settings.dat" For Output As #f
        Write #f, "enableShortCuts", 1
        Write #f, "shortCutUp", rB_SC
        Write #f, "shortCutDown", lB_SC
        Write #f, "start-UP", 0
        Write #f, "runBlckScrn", 0
        Write #f, "shortCutLabel", 1
        Write #f, "languageSelect", 0
    Close #f
    
    rBrightness = rB_SC
    lBrightness = lB_SC
    M_BRIGHTNESS = 128
    SHOW_SHORTCUTS = True
    chckLanguage = 0
End Sub

Public Sub LoadSettings()
    Dim f, Val As Integer
    Dim SettingName, Txt As String
    
    f = FreeFile(0)
    Open "Settings.dat" For Input As #f
    Do Until EOF(f)
        Input #f, SettingName
        Select Case SettingName
            Case "enableShortCuts":
                Input #f, Val
                frmConfig.chEnableSC.Value = Val
            Case "shortCutUp":
                Input #f, Txt
                frmConfig.txtBrightUp.Text = Txt
            Case "shortCutDown":
                Input #f, Txt
                frmConfig.txtBrightDown.Text = Txt
            Case "start-UP":
                Input #f, Val
                frmConfig.chStartUp.Value = Val
            Case "runBlckScrn":
                Input #f, Val
                frmConfig.chOnOff.Value = Val
            Case "shortCutLabel":
                Input #f, Val
                frmConfig.chLabel.Value = Val
            Case "languageSelect":
                Input #f, Val
                frmConfig.cmdLanguage.ListIndex = Val
        End Select
    Loop
    Close #f
End Sub

Public Sub SaveSettings()
    Dim f As Integer
    
    On Error Resume Next: Kill "Settings.dat": On Error GoTo 0
    f = FreeFile(0)
    Open "Settings.dat" For Output As #f
        Write #f, "enableShortCuts", frmConfig.chEnableSC.Value
        Write #f, "shortCutUp", frmConfig.txtBrightUp.Text
        Write #f, "shortCutDown", frmConfig.txtBrightDown.Text
        Write #f, "start-UP", frmConfig.chStartUp.Value
        Write #f, "runBlckScrn", frmConfig.chOnOff.Value
        Write #f, "shortCutLabel", frmConfig.chLabel.Value
        Write #f, "languageSelect", frmConfig.cmdLanguage.ListIndex
    Close #f
End Sub

Private Function fileExistsCheck(f As String) As Boolean
    fileExistsCheck = (PathFileExists(f) <> 0)
End Function
