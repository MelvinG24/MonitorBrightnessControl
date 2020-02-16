VERSION 5.00
Begin VB.UserControl ctrlSlider 
   Alignable       =   -1  'True
   BackStyle       =   0  'Transparent
   ClientHeight    =   3930
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2010
   ScaleHeight     =   3930
   ScaleWidth      =   2010
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   720
      Top             =   960
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   720
      Top             =   480
   End
   Begin VB.Image imgButton 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   0
      MouseIcon       =   "ctrlSlider.ctx":0000
      MousePointer    =   99  'Custom
      Picture         =   "ctrlSlider.ctx":0152
      Stretch         =   -1  'True
      Top             =   0
      Width           =   240
   End
   Begin VB.Image imgButtonPress 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   960
      Picture         =   "ctrlSlider.ctx":0438
      Stretch         =   -1  'True
      Top             =   0
      Width           =   240
   End
   Begin VB.Image imgButtonOver 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   720
      Picture         =   "ctrlSlider.ctx":080B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   240
   End
   Begin VB.Image imgButtonOff 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   480
      Picture         =   "ctrlSlider.ctx":08EA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   240
   End
   Begin VB.Image imgSliderActive 
      Appearance      =   0  'Flat
      Height          =   3120
      Left            =   1200
      Picture         =   "ctrlSlider.ctx":0BD0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   120
   End
   Begin VB.Image imgSlider 
      Appearance      =   0  'Flat
      Height          =   3120
      Left            =   60
      Picture         =   "ctrlSlider.ctx":0C58
      Stretch         =   -1  'True
      Top             =   0
      Width           =   120
   End
End
Attribute VB_Name = "ctrlSlider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Type POINTAPI
    X As Long
    Y As Long
End Type

'Dim CursorPos As POINTAPI
'Dim CurX As Long
Dim CurY As Long

'Dim OffsetX As Long
Dim OffsetY As Long

'Dim ButtonX As Long
Dim ButtonY As Long
Dim ButtonPressed As Boolean

Dim BackGroundY As Long
Dim BackGroundYHigh As Boolean

Dim MaxY As Long
Dim MinY As Long

Private valDirct As Boolean, vsblVal As Boolean
Private intMin As Double, intMax As Double, intSmallChange As Double, intLargeChange As Double, intValue As Double

'Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
'Private Declare Function GetCapture Lib "user32" () As Long

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Event Change()

Public Function GetXCursorPos() As Long
    Dim pt As POINTAPI
    GetCursorPos pt
    GetXCursorPos = pt.X
End Function

Public Function GetYCursorPos() As Long
    Dim pt As POINTAPI
    GetCursorPos pt
    GetYCursorPos = pt.Y
End Function

Public Function ChangeValue() As Double
    Total = MaxY - MinY
    ChangeValue = (imgButton.Top / Total) * 100
End Function

Public Function SetValue(sValue As Double)
    Total = MaxY - MinY
    If sValue >= 0 And sValue <= 100 Then
        imgButton.Top = (sValue / 100) * Total
        
        RaiseEvent Change
    End If
    
    If sValue < 0 Then
        imgButton.Top = MinY
        
        RaiseEvent Change
    End If
    
    If sValue > 100 Then
        imgButton.Top = MaxY
        
        RaiseEvent Change
    End If
End Function

Public Property Get SmallChange() As Double
    SmallChange = intSmallChange
End Property

Public Property Let SmallChange(intVal As Double)
    intSmallChange = intVal
    PropertyChanged "CambioCorto"
End Property

Public Property Get LargeChange() As Double
    LargeChange = intLargeChange
End Property

Public Property Let LargeChange(intVal As Double)
    intLargeChange = intVal
    PropertyChanged "CambioLargo"
End Property

Public Property Get Value() As Double
    Value = ChangeValue
End Property

Public Property Let Value(intVal As Double)
    intValue = intVal
    PropertyChanged "Valor"
    
    Call SetValue(intValue)
End Property

Public Property Get Max() As Double
    Max = intMax
End Property

Public Property Let Max(intVal As Double)
    intMax = intVal
    PropertyChanged "Maximo"
End Property

Public Property Get Min() As Double
    Min = intMin
End Property

Public Property Let Min(intVal As Double)
    intMin = intVal
    PropertyChanged "Minimo"
End Property

Public Property Get ValDownUp() As Boolean
    ValDownUp = valDirct
End Property

Public Property Let ValDownUp(intVal As Boolean)
    valDirct = intVal
    PropertyChanged "Direccion"
End Property

Public Property Get VisibleValue() As Boolean
    VisibleValue = vsblVal
End Property

Public Property Let VisibleValue(intVal As Boolean)
    vsblVal = intVal
    PropertyChanged "VisibleVal"
End Property

Private Sub imgButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        imgButton.Picture = imgButtonPress.Picture
        ButtonPressed = True
        ButtonY = imgButton.Top
        CurY = GetYCursorPos
    End If
    
    RaiseEvent Change
End Sub

Private Sub imgButton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With imgButton
        If (X < 0) Or (Y < 0) Or (X > .Width) Or (Y > .Height) Then
            If ButtonPressed <> True Then
                imgButton.Picture = imgButtonOff.Picture
                Call ReleaseCapture
            End If
        End If
    End With
    
    Dim SetY As Long
    
    If ButtonPressed = True Then
        RaiseEvent Change
        OffsetY = GetYCursorPos
        SetY = ButtonY + (OffsetY - CurY) * 15
        RemDisDown = (imgButton.Top + imgSlider.Height) - imgButton.Height
        RemDisUp = imgButton.Top
        If SetY <= MaxY And SetY >= MinY Then
            imgButton.Move imgButton.Left, SetY
        End If
        If imgButton.Top > MaxY Or SetY > MaxY Then
            imgButton.Top = MaxY
        End If
        If imgButton.Top < MinY Or SetY < MinY Then
            imgButton.Top = MinY
        End If
        vlDirct
    End If
End Sub

Private Sub imgButton_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgButton.Picture = imgButtonOff.Picture
    ButtonPressed = False
End Sub

Private Sub imgSlider_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        BackGroundY = Y - imgButton.Height / 2
        Timer2.Enabled = True
        If BackGroundY > imgButton.Top Then
            BackGroundYHigh = True
        Else
            BackGroundYHigh = False
        End If
    End If
    
    RaiseEvent Change
End Sub

Private Sub imgSlider_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Timer2.Enabled = False
End Sub

Private Sub Timer1_Timer() 'SMALL CHANGE
    Dim RemDisDown As Long 'Remaining Distance (DOWN)
    Dim RemDisUp As Long ' Remaining Distance (UP)
    Dim cValue As Long 'Change Value
    
    cValue = intSmallChange * 10
    
    RemDisDown = (imgButton.Top + imgSlider.Height) - imgButton.Height
    RemDisUp = imgButton.Top
End Sub

Private Sub Timer2_Timer() 'LARGE CHANGE
    Dim cValue As Long 'Change Value
    
    cValue = intLargeChange * 10
    
    If imgButton.Top > BackGroundY And BackGroundYHigh = False Then
        If imgButton.Top <= MinY Then
            imgButton.Top = MinY
            Timer2.Enabled = False
        Else
            imgButton.Top = imgButton.Top - cValue
            RaiseEvent Change
        End If
    ElseIf imgButton.Top < BackGroundY And BackGroundYHigh = True Then
        If imgButton.Top >= MaxY Then
            imgButton.Top = MaxY
            Timer2.Enabled = False
        Else
            imgButton.Top = imgButton + cValue
            RaiseEvent Change
        End If
    End If
End Sub

Private Sub UserControl_Initialize()
    UserControl_Resize
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    vsblVal = PropBag.ReadProperty("VisibleVal", True)
    valDirct = PropBag.ReadProperty("Direccion", True)
    intValue = PropBag.ReadProperty("Valor", 0)
    intMax = PropBag.ReadProperty("Maximo", 32767)
    intMin = PropBag.ReadProperty("Minimo", 0)
    intLargeChange = PropBag.ReadProperty("CambioLargo", 1)
    intSmallChange = PropBag.ReadProperty("CambioCorto", 1)
End Sub

Private Sub UserControl_Resize()
    MaxY = UserControl.Height - imgButton.Height
    MinY = imgSlider.Height
    imgSlider.Height = UserControl.Height
    imgButton.Top = imgSlider.Top
    UserControl.Width = imgButton.Width
    imgButton.Picture = imgButtonOff.Picture
    imgSliderActive.Visible = vsblVal
    vlDirct
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "VisibleVal", vsblVal, True
    PropBag.WriteProperty "Direccion", valDirct, True
    PropBag.WriteProperty "Valor", intValue, 0
    PropBag.WriteProperty "Maximo", intMax, 32767
    PropBag.WriteProperty "Minimo", intMin, 0
    PropBag.WriteProperty "CambioLargo", intLargeChange, 1
    PropBag.WriteProperty "CambioCorto", intSmallChange, 1
End Sub

Private Function vlDirct()
    If valDirct = True Then
        imgSliderActive.Left = imgSlider.Left
        imgSliderActive.Top = imgButton.Top + (imgButton.Height / 2)
        imgSliderActive.Height = imgSlider.Height - (imgButton.Top + (imgButton.Height / 2))
    ElseIf valDirct = False Then
        imgSliderActive.Left = imgSlider.Left
        imgSliderActive.Top = imgSlider.Top
        imgSliderActive.Height = imgButton.Top + (imgButton.Height / 2)
    End If
End Function
