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

'Public Declare Function GetKeyState Lib "user32" ( _
'                ByVal nVirtKey As Long) As Integer

Public Const SPI_GETWORKAREA = 48
Public Const lB_SC As String = "Ctrl + Shift + -"   'Default lower brightness short-cut
Public Const rB_SC As String = "Ctrl + Shift + +"   'Default raise brightness short-cut
'----------------------------------------------------------
' Configuration Variables
'----------------------------------------------------------
Public P_VarBrightnessLevel As Integer              'Brightness level control
Public P_VarChckLanguage As Integer                 'Check program language
Public P_VarChckRunStartUp As Integer               'Check program if run with MS-Windows start-up
Public P_VarChckRunBS As Integer                    'Check black-screen if run after program started
Public P_VarChckSCEnable As Integer                 'Check if short-cut are enable to use or not
Public P_VarChckSCVisible As Integer                'Check if short-cut are visible on black-screen
Public P_VarLwBrightness As String                  'Lower brightness short-cut
Public P_VarRsBrightness As String                  'Raise brightness short-cut
Public L As Integer                                 'Language selected if 0(zero) English, if # Spanish

'----------------------------------------------------------
' Start-up program
'----------------------------------------------------------
Private Sub Main()
    If App.PrevInstance Then Exit Sub
    
    'Default brightness level after run program
    P_VarBrightnessLevel = 128
    
    'Check if settings files exist
    If fileExistsCheck("Settings.dat") Then
        LoadSettings
    Else
        createSettings
    End If
    
    '*******************************
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
' Check/Change language
'----------------------------------------------------------
Public Sub F_L(I As Integer)
    If I = 1 Then
        L = 1255
    Else
        L = 0
    End If
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
    
    P_VarChckLanguage = 0
    chckStartUp 'Function call
    P_VarChckRunBS = 1
    P_VarChckSCEnable = 1
    P_VarChckSCVisible = 1
    P_VarLwBrightness = lB_SC
    P_VarRsBrightness = rB_SC
    L = 0
    
    On Error Resume Next: Kill "Settings.dat": On Error GoTo 0
    f = FreeFile(0)
    Open "Settings.dat" For Output As #f
        Write #f, "varChckLanguage", P_VarChckLanguage
'        Write #f, "varChckRunStartUp", P_VarChckRunStartUp
        Write #f, "varChckRunBS", P_VarChckRunBS
        Write #f, "varChckSCEnable", P_VarChckSCEnable
        Write #f, "varChckSCVisible", P_VarChckSCVisible
        Write #f, "varLwBrightness", P_VarLwBrightness
        Write #f, "varRsBrightness", P_VarRsBrightness
    Close #f
End Sub

Public Sub LoadSettings()
    Dim f, Val As Integer
    Dim SettingName, Txt As String
    
    chckStartUp 'Function call
    
    f = FreeFile(0)
    Open "Settings.dat" For Input As #f
    Do Until EOF(f)
        Input #f, SettingName
        Select Case SettingName
            Case "varChckLanguage":
                Input #f, Val
                P_VarChckLanguage = Val
'            Case "varChckRunStartUp":
'                Input #f, Val
'                P_VarChckRunStartUp = Val
            Case "varChckRunBS":
                Input #f, Val
                P_VarChckRunBS = Val
            Case "varChckSCEnable":
                Input #f, Val
                P_VarChckSCEnable = Val
            Case "varChckSCVisible":
                Input #f, Val
                P_VarChckSCVisible = Val
            Case "varLwBrightness":
                Input #f, Txt
                P_VarLwBrightness = Txt
            Case "varRsBrightness":
                Input #f, Txt
                P_VarRsBrightness = Txt
        End Select
    Loop
    Close #f
    
    F_L P_VarChckLanguage
End Sub

Public Sub SaveSettings()
    Dim f As Integer
    
    With frmConfig
        P_VarChckLanguage = .cmdLanguage.ListIndex
        chckStartUp 'Function call
        P_VarChckRunBS = .ChckRunBS.Value
        P_VarChckSCEnable = .ChckSCEnable.Value
        P_VarChckSCVisible = .ChckSCVisible.Value
        P_VarLwBrightness = .txtBrightDown.Text
        P_VarRsBrightness = .txtBrightUp.Text
    End With
    
    On Error Resume Next: Kill "Settings.dat": On Error GoTo 0
    f = FreeFile(0)
    Open "Settings.dat" For Output As #f
        Write #f, "varChckLanguage", P_VarChckLanguage
'            Write #f, "varChckRunStartUp", .ChckRunStartUp.Value
        Write #f, "varChckRunBS", P_VarChckRunBS
        Write #f, "varChckSCEnable", P_VarChckSCEnable
        Write #f, "varChckSCVisible", P_VarChckSCVisible
        Write #f, "varLwBrightness", P_VarLwBrightness
        Write #f, "varRsBrightness", P_VarRsBrightness
    Close #f
End Sub

Private Function fileExistsCheck(f As String) As Boolean
    fileExistsCheck = (PathFileExists(f) <> 0)
End Function

'----------------------------------------------------------
' Check if program start-up with MS-Windows
'----------------------------------------------------------
Private Function chckStartUp()
    m_IgnoreEvents = True
    If StartUp(App.EXEName) Then
        P_VarChckRunStartUp = 1
    Else
        P_VarChckRunStartUp = 0
'        P_VarChckRunBS = 0
    End If
    m_IgnoreEvents = False
End Function
