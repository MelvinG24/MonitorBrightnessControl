Attribute VB_Name = "basRegKey"
' Original code by Neil Crosby of vb-helper.com
' Original code URL http://vb-helper.com/howto_run_at_startup.html
Option Explicit

Public m_IgnoreEvents As Boolean

Private Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" ( _
                ByVal hKey As Long, _
                ByVal lpSubKey As String, _
                ByVal Reserved As Long, _
                ByVal lpClass As String, _
                ByVal dwOptions As Long, _
                ByVal samDesired As Long, _
                ByVal lpSecurityAttributes As Long, _
                phkResult As Long, _
                lpdwDisposition As Long) As Long

Private Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" ( _
                ByVal hKey As Long, _
                ByVal lpValueName As String, _
                ByVal Reserved As Long, _
                ByVal dwType As Long, _
                lpData As Any, _
                ByVal cbData As Long) As Long

Private Declare Function RegDeleteValue Lib "advapi32" Alias "RegDeleteValueA" ( _
                ByVal hKey As Long, _
                ByVal lpValueName As String) As Long

Private Declare Function RegCloseKey Lib "advapi32" ( _
                ByVal hKey As Long) As Long

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" ( _
                ByVal hKey As Long, _
                ByVal lpSubKey As String, _
                ByVal ulOptions As Long, _
                ByVal samDesired As Long, _
                phkResult As Long) As Long

Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" ( _
                ByVal hKey As Long, _
                ByVal lpValueName As String, _
                ByVal lpReserved As Long, _
                lpType As Long, _
                lpData As Any, _
                lpcbData As Long) As Long

Private Const READ_CONTROL = &H20000
Private Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Private Const SYNCHRONIZE = &H100000

Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)

Private Const ERROR_SUCCESS = 0&

Private Const HKEY_CURRENT_USER = &H80000001
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Private Const KEY_SET_VALUE = &H2
Private Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))

Private Const REG_SZ = 1

Public Function StartUp(ByVal appName As String) As Boolean
    Dim hKey As Long
    Dim valueType As Long
    
    If RegOpenKeyEx(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", 0, KEY_READ, hKey) = ERROR_SUCCESS Then
        StartUp = (RegQueryValueEx(hKey, appName, ByVal 0&, valueType, ByVal 0&, ByVal 0&) = ERROR_SUCCESS)
        RegCloseKey hKey
    Else
        StartUp = False
    End If
End Function

Public Sub SetRunAtStartUp(ByVal appName As String, ByVal appPath As String, Optional ByVal runAtStartUp As Boolean)
    Dim hKey As Long
    Dim keyValue As String
    Dim Status As Long
    
    On Error GoTo SetStartUpError
    
    If RegCreateKeyEx(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", ByVal 0&, ByVal 0&, ByVal 0&, KEY_WRITE, ByVal 0&, hKey, ByVal 0&) <> ERROR_SUCCESS Then
        MsgBox "Error " & Err.Number & " opening key" & vbCrLf & Err.Description
    End If
    
    If runAtStartUp Then
        keyValue = appPath & "\" & appName & ".exe" & vbNullChar
        Status = RegSetValueEx(hKey, App.EXEName, 0, REG_SZ, ByVal keyValue, Len(keyValue))
        If Status <> ERROR_SUCCESS Then
            MsgBox "Error " & Err.Number & " settings key" & vbCrLf & Err.Description
        End If
    Else
        RegDeleteValue hKey, appName
    End If
    
    RegCloseKey hKey
    Exit Sub
    
SetStartUpError:
    MsgBox Err.Number & " " & Err.Description
    Exit Sub
End Sub
