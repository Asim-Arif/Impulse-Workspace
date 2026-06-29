Attribute VB_Name = "PopUpMenu"

Private Declare Function CreatePopupMenu Lib "User32" () As Long
Private Declare Function AppendMenu Lib "User32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Private Declare Function TrackPopupMenu Lib "User32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal x As Long, ByVal y As Long, ByVal nReserved As Long, ByVal hwnd As Long, ByVal nIgnored As Long) As Long
Private Declare Function DestroyMenu Lib "User32" (ByVal hMenu As Long) As Long
Public Declare Function GetCursorPos Lib "User32" (lpPoint As POINTAPI) As Long
Public Declare Function GetForegroundWindow Lib "User32" () As Long
Public Declare Function GetClassName Lib "User32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

Private Declare Function GetMenuString Lib "User32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Type POINTAPI
        x As Long
        y As Long
End Type
Private Const MF_ENABLED = &H0&
Private Const MF_SEPARATOR = &H800&
Private Const MF_STRING = &H0&
Private Const TPM_LEFTALIGN = &H0&
Private Const TPM_RIGHTALIGN = &H8&
Private Const TPM_RIGHTBUTTON = &H2&
Private Const TPM_NONOTIFY = &H80&
Private Const TPM_RETURNCMD = &H100&
Private Const MF_DISABLED = &H2&
Private Const MF_GRAYED = &H1&


Dim Point As POINTAPI

Public Function ShowPopUpMenu(lLeft As Long, lTop As Long, ParamArray Param()) As String

    Dim iMenu As Long
    Dim hMenu As Long
    Dim nMenus As Long
    Dim str As String
    str = Space$(255)
    Dim P As POINTAPI
    
    GetCursorPos P
    
    hMenu = CreatePopupMenu()
    
    nMenus = 1 + UBound(Param)
    
    For iMenu = 1 To nMenus
        If Trim$(CStr(Param(iMenu - 1))) = "-" Then
            AppendMenu hMenu, MF_SEPARATOR, iMenu, ""
        Else
            AppendMenu hMenu, MF_STRING + MF_ENABLED, iMenu, CStr(Param(iMenu - 1))
        End If
    Next iMenu
    
    'iMenu = TrackPopupMenu(hMenu, TPM_RIGHTBUTTON + TPM_LEFTALIGN + TPM_NONOTIFY + TPM_RETURNCMD, p.X, p.Y, 0, GetForegroundWindow(), 0)
    'iMenu = TrackPopupMenu(hMenu, TPM_RIGHTBUTTON + TPM_LEFTALIGN + TPM_NONOTIFY + TPM_RETURNCMD, p.X, p.Y, 0, MainForm.hwnd, 0)
    iMenu = TrackPopupMenu(hMenu, TPM_RIGHTBUTTON + TPM_LEFTALIGN + TPM_NONOTIFY + TPM_RETURNCMD, lLeft / 15, lTop / 15, 0, MainForm.hwnd, 0)
    Dim RetLen As Integer
    RetLen = GetMenuString(hMenu, iMenu, str, 255, 0)
    
    DestroyMenu hMenu
    
    ShowPopUpMenu = Left(str, RetLen)
    
End Function


Public Function ShowPopUpMenuAdv(lLeft As Long, lTop As Long, ParamArray Param()) As String
    
    
    On Error GoTo err
    
    Dim iMenu As Long
    Dim hMenu As Long
    Dim nMenus As Long
    Dim str As String
    str = Space$(255)
    Dim P As POINTAPI
    
    GetCursorPos P
    
    hMenu = CreatePopupMenu()
    
    nMenus = 1 + UBound(Param)
    
    For iMenu = 1 To nMenus
        If Trim$(CStr(Param(iMenu - 1))) = "-" Then
            AppendMenu hMenu, MF_SEPARATOR, iMenu, ""
        Else
            'AppendMenu hMenu, MF_STRING + MF_ENABLED, iMenu, CStr(Param(iMenu - 1))
            Call SetMenu(CStr(Param(iMenu - 1)), hMenu, iMenu)
        End If
    Next iMenu
    
    'iMenu = TrackPopupMenu(hMenu, TPM_RIGHTBUTTON + TPM_LEFTALIGN + TPM_NONOTIFY + TPM_RETURNCMD, p.X, p.Y, 0, GetForegroundWindow(), 0)
    'iMenu = TrackPopupMenu(hMenu, TPM_RIGHTBUTTON + TPM_LEFTALIGN + TPM_NONOTIFY + TPM_RETURNCMD, p.X, p.Y, 0, MainForm.hwnd, 0)
    If App.Title = "HicoExport" Then
        iMenu = TrackPopupMenu(hMenu, TPM_RIGHTBUTTON + TPM_LEFTALIGN + TPM_NONOTIFY + TPM_RETURNCMD, lLeft / 15, lTop / 15, 0, MainForm.hwnd, 0)
    Else
        iMenu = TrackPopupMenu(hMenu, TPM_RIGHTBUTTON + TPM_LEFTALIGN + TPM_NONOTIFY + TPM_RETURNCMD, lLeft / 15, lTop / 15, 0, MainForm.hwnd, 0)
    End If
    
    Dim RetLen As Integer
    RetLen = GetMenuString(hMenu, iMenu, str, 255, 0)
    
    DestroyMenu hMenu
    
    ShowPopUpMenuAdv = Left(str, RetLen)
    
    Exit Function
err:
    MsgBox err.Description
End Function

Private Sub SetMenu(strtemp As String, hMenu As Long, iMenu As Long)
    
    
    Dim FieldName As String
    Dim MenuName As String
    
    MenuName = Left(strtemp, InStr(strtemp, "|||") - 1)
    
    FieldName = Mid(strtemp, InStr(strtemp, "|||") + 3)
    
    If rsUserSettings.Fields(FieldName) Then
        AppendMenu hMenu, MF_STRING + MF_ENABLED, iMenu, MenuName
    Else
        AppendMenu hMenu, MF_STRING + MF_GRAYED, iMenu, MenuName
    End If
    
End Sub

Public Function ShowPopUpMenu1(ParamArray Param()) As String

    Dim iMenu As Long
    Dim hMenu As Long
    Dim nMenus As Long
    Dim str As String
    str = Space$(255)
    Dim P As POINTAPI
    
    GetCursorPos P
    
    hMenu = CreatePopupMenu()
    
    nMenus = 1 + UBound(Param)
    
    For iMenu = 1 To nMenus
        If Trim$(CStr(Param(iMenu - 1))) = "-" Then
            AppendMenu hMenu, MF_SEPARATOR, iMenu, ""
        ElseIf Trim$(CStr(Param(iMenu - 1))) <> "" Then
            AppendMenu hMenu, MF_STRING + MF_ENABLED, iMenu, CStr(Param(iMenu - 1))
        End If
    Next iMenu
    
    'iMenu = TrackPopupMenu(hMenu, TPM_RIGHTBUTTON + TPM_LEFTALIGN + TPM_NONOTIFY + TPM_RETURNCMD, p.X, p.Y, 0, GetForegroundWindow(), 0)
    'iMenu = TrackPopupMenu(hMenu, TPM_RIGHTBUTTON + TPM_LEFTALIGN + TPM_NONOTIFY + TPM_RETURNCMD, p.X, p.Y, 0, MainForm.hwnd, 0)
    iMenu = TrackPopupMenu(hMenu, TPM_LEFTALIGN + TPM_NONOTIFY + TPM_RETURNCMD, P.x, P.y, 0, MainForm.hwnd, 0)
    Dim RetLen As Integer
    RetLen = GetMenuString(hMenu, iMenu, str, 255, 0)
    
    DestroyMenu hMenu
    
    ShowPopUpMenu1 = Left(str, RetLen)
    
End Function
