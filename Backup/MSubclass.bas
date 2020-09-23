Attribute VB_Name = "MSubclass"
Public Const GWL_WNDPROC = -4
Public Const WM_QUERYENDSESSION = &H11
Public Const WM_CANCELMODE = &H1F

Type POINTAPI
        x As Long
        y As Long
End Type

Type MSG
    hwnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type

Public ShutDownMessage As MSG

Global lpPrevWndProc As Long
Global gHW As Long

Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function ExitWindows Lib "user32" (ByVal dwReserved As Long, ByVal uReturnCode As Long) As Long
Declare Function GetMessage Lib "user32" Alias "GetMessageA" (lpMsg As MSG, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long

Public Sub Hook()
    lpPrevWndProc = SetWindowLong(gHW, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Public Sub Unhook()
    Dim temp As Long
    temp = SetWindowLong(gHW, GWL_WNDPROC, lpPrevWndProc)
End Sub

Function WindowProc(ByVal hw As Long, ByVal uMsg As _
Long, ByVal wParam As Long, ByVal lParam As Long) As _
Long
Dim a As Long
    If uMsg = WM_QUERYENDSESSION Then
        If vbYes = MsgBox("Do You want to backup now ?", vbYesNo + vbQuestion, "Backup") Then
            DoBackup
        End If
    End If
    WindowProc = CallWindowProc(lpPrevWndProc, hw, _
    uMsg, wParam, lParam)
End Function
