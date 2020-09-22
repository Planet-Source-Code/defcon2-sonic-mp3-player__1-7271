Attribute VB_Name = "MODUL_Main"
'Declares
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
Public Declare Function WritePrivateProfileString% Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName$, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName$)
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long


Public Const SWP_NOMOVE = 2
Public Const SWP_NOSIZE = 1
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2




'Windows 2000

Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Public Declare Function UpdateLayeredWindow Lib "user32" (ByVal hWnd As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hdcSrc As Long, pptSrc As Any, crKey As Long, ByVal pblend As Long, ByVal dwFlags As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Const GWL_EXSTYLE = (-20)
Public Const LWA_COLORKEY = &H1
Public Const LWA_ALPHA = &H2
Public Const ULW_COLORKEY = &H1
Public Const ULW_ALPHA = &H2
Public Const ULW_OPAQUE = &H4
Public Const WS_EX_LAYERED = &H80000


Public Type POINTAPI

    X As Long
    Y As Long
    
End Type

'Ende Windows 2000


' ## Browse for Folders

Public Type BROWSEINFOTYPE
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Public Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBROWSEINFOTYPE As BROWSEINFOTYPE) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)

Public Const WM_USER = &H400

Public Const BFFM_SETSELECTIONA As Long = (WM_USER + 102)
Public Const BFFM_SETSELECTIONW As Long = (WM_USER + 103)
Public Declare Function LocalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal uBytes As Long) As Long
Public Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Public Const LPTR = (&H0 Or &H40)

' Ende Browse for Folders



'Subs

Public Function GetX() As Long

    Dim n As POINTAPI
    GetCursorPos n
    GetX = n.X
    
End Function


Public Function GetY() As Long

    Dim n As POINTAPI
    GetCursorPos n
    GetY = n.Y
    
End Function

Sub AlwaysOnTop(frmID As Form, OnTop As Integer)
        
    If OnTop Then
        OnTop = SetWindowPos(frmID.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
    Else
        OnTop = SetWindowPos(frmID.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
    End If
    
End Sub



Public Function isTransparent(ByVal hWnd As Long) As Boolean
        
    On Error Resume Next
        
    Dim Msg As Long
    
        Msg = GetWindowLong(hWnd, GWL_EXSTYLE)

    If (Msg And WS_EX_LAYERED) = WS_EX_LAYERED Then
        isTransparent = True
    Else
        isTransparent = False
    End If


    If Err Then
        isTransparent = False
    End If

End Function

Public Function EnableTransparanty(ByVal hWnd As Long, Perc As Integer) As Long
        
    On Error Resume Next
    
    Dim Msg As Long

    If Perc < 0 Or Perc > 255 Then
        EnableTransparanty = 1
    Else

        Msg = GetWindowLong(hWnd, GWL_EXSTYLE)
        Msg = Msg Or WS_EX_LAYERED
    
        SetWindowLong hWnd, GWL_EXSTYLE, Msg
        SetLayeredWindowAttributes hWnd, 0, Perc, LWA_ALPHA
        EnableTransparanty = 0
    End If

    If Err Then
        EnableTransparanty = 2
    End If
    
End Function

Public Function DisableTransparanty(ByVal hWnd As Long) As Long

    On Error Resume Next
    
    Dim Msg As Long
           
        Msg = GetWindowLong(hWnd, GWL_EXSTYLE)
        Msg = Msg And Not WS_EX_LAYERED

        SetWindowLong hWnd, GWL_EXSTYLE, Msg

        SetLayeredWindowAttributes hWnd, 0, 0, LWA_ALPHA
        DisableTransparanty = 0

    If Err Then
        DisableTransparanty = 2
    End If
   
End Function


Public Function GetFromInI(strSectionHeader As String, strVariableName As String, strFileName As String) As String

    Dim strReturn As String
    
    strReturn = String(255, Chr(0))
    GetFromInI = Left$(strReturn, GetPrivateProfileString(strSectionHeader, ByVal strVariableName, "", strReturn, Len(strReturn), strFileName))
    
End Function

Sub Pause(Dur)

    T = Timer
    Do
    DoEvents
    Loop Until Timer >= T + Dur
    
End Sub

Public Function FileExists(strPath As String) As Integer

    FileExists = Not (Dir(strPath) = "")
    
End Function

Public Function BrowseCallbackProcStr(ByVal hWnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long
If uMsg = 1 Then
    Call SendMessage(hWnd, BFFM_SETSELECTIONA, True, ByVal lpData)
End If
End Function

Public Function FunctionPointer(FunctionAddress As Long) As Long
FunctionPointer = FunctionAddress
End Function

Sub ShowMsgBox(MsgCaption As String, MsgText As String)

    frmMsgbox.lblMsgBoxCaption.Caption = MsgCaption
    frmMsgbox.lblMsgBoxText.Caption = MsgText
    
    AlwaysOnTop frmMsgbox, 1
    frmMsgbox.Show
    
End Sub

Function KommaToSign(strin As String)

Let inptxt$ = strin
Let lenth% = Len(inptxt$)

Do While numspc% <= lenth%
DoEvents
Let numspc% = numspc% + 1
Let nextchr$ = Mid$(inptxt$, numspc%, 1)
Let nextchrr$ = Mid$(inptxt$, numspc%, 2)
If nextchrr$ = "ae" Then Let nextchrr$ = "": Let newsent$ = newsent$ + nextchrr$: Let crapp% = 2: GoTo Greed
If nextchrr$ = "AE" Then Let nextchrr$ = "": Let newsent$ = newsent$ + nextchrr$: Let crapp% = 2: GoTo Greed
If nextchrr$ = "oe" Then Let nextchrr$ = "": Let newsent$ = newsent$ + nextchrr$: Let crapp% = 2: GoTo Greed
If nextchrr$ = "OE" Then Let nextchrr$ = "": Let newsent$ = newsent$ + nextchrr$: Let crapp% = 2: GoTo Greed
If crapp% > 0 Then GoTo Greed

If nextchr$ = "," Then Let nextchr$ = "#"

Let newsent$ = newsent$ + nextchr$

Greed:
If crapp% > 0 Then Let crapp% = crapp% - 1
DoEvents
Loop
KommaToSign = newsent$

End Function

Function SignToKomma(strin As String)

Let inptxt$ = strin
Let lenth% = Len(inptxt$)

Do While numspc% <= lenth%
DoEvents
Let numspc% = numspc% + 1
Let nextchr$ = Mid$(inptxt$, numspc%, 1)
Let nextchrr$ = Mid$(inptxt$, numspc%, 2)
If nextchrr$ = "ae" Then Let nextchrr$ = "": Let newsent$ = newsent$ + nextchrr$: Let crapp% = 2: GoTo Greed
If nextchrr$ = "AE" Then Let nextchrr$ = "": Let newsent$ = newsent$ + nextchrr$: Let crapp% = 2: GoTo Greed
If nextchrr$ = "oe" Then Let nextchrr$ = "": Let newsent$ = newsent$ + nextchrr$: Let crapp% = 2: GoTo Greed
If nextchrr$ = "OE" Then Let nextchrr$ = "": Let newsent$ = newsent$ + nextchrr$: Let crapp% = 2: GoTo Greed
If crapp% > 0 Then GoTo Greed

If nextchr$ = "#" Then Let nextchr$ = ","

Let newsent$ = newsent$ + nextchr$

Greed:
If crapp% > 0 Then Let crapp% = crapp% - 1
DoEvents
Loop
SignToKomma = newsent$

End Function

