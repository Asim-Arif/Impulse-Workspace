Attribute VB_Name = "FormEffects"

Option Explicit

Public Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type
Public Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean
Public Const ICC_USEREX_CLASSES = &H200


Private Declare Function GetWindowLong Lib "User32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Declare Function SetLayeredWindowAttributes Lib "User32" (ByVal hwnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_ALPHA = &H2&
Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Public Declare Function SetWindowRgn Lib "User32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long



Public Function TransparentMe(f As Form, Optional DefaultTransparency As Boolean)
    
    If DefaultTransparency Then
        Call SetLayeredWindowAttributes(f.hwnd, 0, 254, LWA_ALPHA)
    Else
        Call SetLayeredWindowAttributes(f.hwnd, 0, 1, LWA_ALPHA)
    End If
    
End Function

Public Function FadeMe(f As Form, Optional Showing As Boolean = False)
Dim LEVEL As Byte
Dim i As Long
For i = 1 To 256 Step 16
    If i = 17 Then i = 16
    If Showing Then
       LEVEL = i - 1
    Else
       LEVEL = 256 - i
    End If
    
    Call SetWindowLong(f.hwnd, GWL_EXSTYLE, GetWindowLong(f.hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED)
    Call SetLayeredWindowAttributes(f.hwnd, 0, LEVEL, LWA_ALPHA)
    On Error Resume Next
        'If I Mod 96 Then f.Refresh
        If i = 1 Then f.Refresh
    On Error GoTo 0
    Sleep 50
Next
    Call SetLayeredWindowAttributes(f.hwnd, 0, LEVEL, LWA_ALPHA)
    'If Not DontRestore Then Call TransparentMe(f, True)
End Function


