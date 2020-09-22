Attribute VB_Name = "MODUL_Spriter"
Declare Function CreateRectRgn Lib "GDI32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function ReleaseCapture Lib "user32" () As Long
Declare Function GetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long) As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long

Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long


Public Const WM_MOVE = &H3
Public Const WM_DESTROY = &H2
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2
'Public Const GWL_WNDPROC = (-4)


Option Explicit

Private m_lOriginalProc As Long
Private m_hwndOriginal As Long
Private m_ctrlClient As Object
Private m_bIsHook As Boolean

Public Sub IndirectMoveWindow(ByVal hwndClient As Long)

     Call ReleaseCapture
     Call SendMessage(hwndClient, WM_NCLBUTTONDOWN, HTCAPTION, 0&)

End Sub

'Gwyshell
'The Window Proc()
Public Function ControlWndProc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    Select Case Msg
    
        Case WM_MOVE
            m_ctrlClient.FireWindowMove wParam, lParam
            Debug.Print "Window Moved", wParam, ":", lParam
        Case WM_DESTROY
            'ControlWndProc = CallWindowProc(m_lOriginalProc, hWnd, Msg, wParam, lParam)
            ReleaseControlHook
            Exit Function
        
    End Select
    
    'Call Origianl WinProc
    ControlWndProc = CallWindowProc(m_lOriginalProc, hWnd, Msg, wParam, lParam)

End Function

Public Sub SetControlHook(ByVal ctrl As Object)

'For Debug and Prevent From Crashing VB
#If Debug_Win32 Then
'    Exit Sub
#End If

    'Release before Hooking
    If m_bIsHook Then
        ReleaseControlHook
    End If
    Set m_ctrlClient = ctrl
    m_hwndOriginal = m_ctrlClient.hWnd
    m_lOriginalProc = SetWindowLong(m_hwndOriginal, GWL_WNDPROC, AddressOf ControlWndProc)
    'Debug.Print "Hook Window = " & m_hwndOriginal, m_bIsHook
    m_bIsHook = True

End Sub

Public Sub ReleaseControlHook()

    If m_bIsHook Then
        If GetWindowLong(m_hwndOriginal, GWL_WNDPROC) <> m_lOriginalProc Then
            SetWindowLong m_hwndOriginal, GWL_WNDPROC, m_lOriginalProc
        End If
        'Debug.Print "DeHook Window = " & m_hwndOriginal, m_bIsHook
        m_bIsHook = False
        Set m_ctrlClient = Nothing
    End If

End Sub
