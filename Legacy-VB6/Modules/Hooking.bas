Attribute VB_Name = "Hooking"
Option Explicit


Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Const WM_ACTIVATEAPP = &H1C
Public Const GWL_WNDPROC = -4

Global lpPrevWndProc As Long
Global gHW As Long

'''''' Red=255,Green=211,Blue=210

Public Sub Hook()

    'Establish a hook to capture messages to this window.
    Exit Sub
    If Command() = "FromVBIDE" Then Exit Sub
    
    lpPrevWndProc = SetWindowLong(gHW, GWL_WNDPROC, _
        AddressOf WindowProc)
    
End Sub

Public Sub Unhook()
    
    Exit Sub
    If Command() = "FromVBIDE" Then Exit Sub
    
    Dim Temp As Long

    'Reset the message handler for this window.
    Temp = SetWindowLong(gHW, GWL_WNDPROC, lpPrevWndProc)
    
End Sub

Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, _
   ByVal wParam As Long, ByVal lParam As Long) As Long
   
   Dim i As Integer
    'Check for the ActivateApp message.
    If uMsg = WM_ACTIVATEAPP Then
        'Check to see if Activating the application.
        If wParam <> 0 Then
            'Check to see if application is minimized.
            'If MainForm.WindowState <> vbMinimized Then
            
            frmAway.Show 1
            
            'End If
        'Else
            'Application is DeActivating so restore normal colors.
        End If
    
    End If

    'Pass message on to the original window message handler.
    WindowProc = CallWindowProc(lpPrevWndProc, hw, uMsg, wParam, _
      lParam)
End Function
