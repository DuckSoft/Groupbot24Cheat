Attribute VB_Name = "Module1"
Option Explicit

Declare Function RegisterHotKey Lib "user32" (ByVal hWnd As Long, ByVal id As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
Declare Function UnregisterHotKey Lib "user32" (ByVal hWnd As Long, ByVal id As Long) As Long
Public Const WM_HOTKEY = &H312
Public RHK_HOME_ID, RHK_END_ID As Long

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const GWL_WNDPROC = (-4)
Public preWinProc As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Public Sub Sleep(ByVal msec As Long)
    Dim iTick As Long
    iTick = GetTickCount
    While GetTickCount - iTick < msec
        DoEvents
    Wend
End Sub
Public Function wndproc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If Msg = WM_HOTKEY Then
        Select Case (wParam)
        Case 0:
            Sleep 500
            SendKeys "¿ªÊ¼24µã^{Enter}", True
        End Select
    End If

    wndproc = CallWindowProc(preWinProc, hWnd, Msg, wParam, lParam)
End Function
