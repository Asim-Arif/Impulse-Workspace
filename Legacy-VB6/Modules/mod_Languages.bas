Attribute VB_Name = "mod_Languages"
Option Explicit
'api's to adjust the keyboardlayout
Private Declare Function GetKeyboardLayoutName Lib "user32" Alias "GetKeyboardLayoutNameA" (ByVal pwszKLID As String) As Long
Private Declare Function LoadKeyboardLayout Lib "user32" Alias "LoadKeyboardLayoutA" (ByVal pwszKLID As String, ByVal flags As Long) As Long
Private Declare Function GetLastError Lib "kernel32" () As Long

Const KLF_ACTIVATE = &H1

'Languages

Const Lang_EN_USA As String = "00000409"

Const Lang_PK_URD As String = "A0000420"

Public Function SetUrduLanguage() As Boolean
    SetUrduLanguage = SetKbLayout(Lang_PK_URD)
End Function

Public Function SetEnglishLanguage() As Boolean
    SetEnglishLanguage = SetKbLayout(Lang_EN_USA)
End Function

Public Function SetKbLayout(strLocaleId As String) As Boolean
    
    'Changes the KeyboardLayout
    'Returns TRUE when the KeyboardLayout was adjusted properly, FALSE otherwise
    'If the KeyboardLayout isn't installed, this function will install it for you
    On Error Resume Next
    Dim strLocID As String 'used to retrieve current KeyboardLayout

    'create a buffer
    strLocID = String(9, 0)
    'retrieve the current KeyboardLayout
    GetKeyboardLayoutName strLocID
    'Check whether the current KeyboardLayout and the
    'new one are the same
    If strLocID = (strLocaleId & Chr(0)) Then
        'If they're the same, we return immediately
        SetKbLayout = True
        Exit Function
    Else
        'create buffer
        strLocID = String(9, 0)
        
        'load and activate the layout for the current thread
        strLocID = LoadKeyboardLayout((strLocaleId & Chr(0)), KLF_ACTIVATE)

 
    End If
    
    'Test success

    GetKeyboardLayoutName strLocID
    If strLocID = (strLocaleId) Then SetKbLayout = True
    
End Function

