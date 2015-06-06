Attribute VB_Name = "Module1"
    Option Explicit
     
    Private Const GWL_EXSTYLE = (-20)
    Private Const WS_EX_LAYERED = &H80000
    Private Const LWA_ALPHA = &H2
     
    Private Declare Function GetWindowLong Lib "user32" _
    Alias "GetWindowLongA" (ByVal hWnd As Long, _
    ByVal nIndex As Long) As Long
     
    Private Declare Function SetWindowLong Lib "user32" _
    Alias "SetWindowLongA" (ByVal hWnd As Long, _
    ByVal nIndex As Long, ByVal dwNewLong As Long) _
    As Long
     
    Private Declare Function SetLayeredWindowAttributes Lib _
    "user32" (ByVal hWnd As Long, ByVal crKey As Long, _
    ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
     
    Public Function TransForm(Form As Form, TransLevel As Byte) As Boolean
    SetWindowLong Form.hWnd, GWL_EXSTYLE, WS_EX_LAYERED
    SetLayeredWindowAttributes Form.hWnd, 0, TransLevel, LWA_ALPHA
    TransForm = Err.LastDllError = 0
    End Function


