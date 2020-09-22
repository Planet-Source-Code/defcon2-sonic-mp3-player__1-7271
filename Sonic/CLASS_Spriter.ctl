VERSION 5.00
Begin VB.UserControl CLASS_Spriter 
   Appearance      =   0  '2D
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BackStyle       =   0  'Transparent
   ClientHeight    =   2025
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3165
   ScaleHeight     =   2025
   ScaleWidth      =   3165
End
Attribute VB_Name = "CLASS_Spriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' Gwyshell Transparent UserControl
' --------
' Here I provide a simple way to create a non-rectangle window region to
' make parent of the control non-rectangle.
' -----------------------------------------
' Non-Rectangle Form Control Control
'

Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'Event Move(x As Single, y As Single)

Private m_bEnableMoveWindow As Boolean

Private Sub UserControl_Initialize()

    'Hook the Control Window Proc to Get the Specail Message
    'MSpriter.SetControlHook Me
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Me.MaskColor = PropBag.ReadProperty("MaskColor", &H8000000F)
    Set Me.MaskPicture = PropBag.ReadProperty("MaskPicture", Nothing)
    Set Me.MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    Me.MousePointer = PropBag.ReadProperty("MousePointer", MousePointerConstants.vbArrow)
    Me.EnableMoveWindow = PropBag.ReadProperty("EnableMoveWindow", False)

End Sub

Private Sub UserControl_Terminate()
    'MSpriter.ReleaseControlHook
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    PropBag.WriteProperty "MaskColor", Me.MaskColor, &H8000000F
    PropBag.WriteProperty "MaskPicture", Me.MaskPicture, Nothing
    PropBag.WriteProperty "MouseIcon", Me.MouseIcon, Nothing
    PropBag.WriteProperty "MousePointer", Me.MousePointer, MousePointerConstants.vbArrow
    PropBag.WriteProperty "EnableMoveWindow", Me.EnableMoveWindow, False

End Sub

Public Property Get hWnd() As OLE_HANDLE
    hWnd = UserControl.hWnd
End Property

Public Property Get hDC() As OLE_HANDLE
    hDC = UserControl.hDC
End Property

Public Property Get hRgn() As OLE_HANDLE
    
    hRgn = CreateRectRgn(0, 0, 1, 1)
    GetWindowRgn Me.hWnd, hRgn

End Property

Public Property Get MaskPicture() As Picture
    Set MaskPicture = UserControl.MaskPicture
End Property

Public Property Set MaskPicture(ByVal picNew As Picture)
        
    Set UserControl.MaskPicture = picNew
    'Put the Refresh() code before the Set Picture Property will
    'have better effection
    Me.Refresh
    Set UserControl.Picture = picNew
    
    
    PropertyChanged "MaskPicture"

End Property

Public Property Get MousePointer() As MousePointerConstants
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal iMousePointerNew As MousePointerConstants)
    UserControl.MousePointer = iMousePointerNew
    PropertyChanged "MousePointer"
End Property

Public Property Get MouseIcon() As Picture
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

Public Property Get MaskColor() As OLE_COLOR
    MaskColor = UserControl.MaskColor
End Property

Public Property Let MaskColor(ByVal clrMaskColor As OLE_COLOR)
    UserControl.MaskColor = clrMaskColor
    Me.Refresh
    PropertyChanged "MaskColor"
End Property

'To identify if we want to move the parent when dragging the control
Public Property Get EnableMoveWindow() As Boolean
    EnableMoveWindow = m_bEnableMoveWindow
End Property

Public Property Let EnableMoveWindow(ByVal bMove As Boolean)
    m_bEnableMoveWindow = bMove
    PropertyChanged "EnableMoveWindow"
End Property


Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    On Local Error Resume Next

    If Me.EnableMoveWindow Then
        If Button = vbLeftButton Then
            MSpriter.IndirectMoveWindow UserControl.Parent.hWnd
        End If
    End If
    
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

'Refresh
Public Sub Refresh()

    'On Local Error Resume Next
    
    Dim hRgnNormal As Long

    With UserControl
            
        
        If .MaskPicture = 0 Then
            hRgnNormal = CreateRectRgn(0, 0, .ScaleX(.Width), .ScaleY(.Height))
            SetWindowRgn .Extender.Container.hWnd, hRgnNormal, True
        Else

            .Size .ScaleX(.MaskPicture.Width), .ScaleY(.MaskPicture.Height)
            .Extender.Container.Width = .Width
            .Extender.Container.Height = .Height
            .Extender.Move 0, 0
            
            'Gwyshell
            'Let the system have time to finish the special regions created
            DoEvents
            
            'Set New Regions
            AssignNewRegion .Extender.Container.hWnd

            
            If Err Then
                MsgBox "The Container don't support the Width or Height Property!"
            End If
            
        End If
                
    End With

End Sub

'Public Sub FireWindowMove(ByVal lX As Long, ByVal lY As Long)
'
'    With UserControl
'
'        RaiseEvent Move(.ScaleX(lX), .ScaleY(lY))
'
'    End With
'
'End Sub

'Create the Mask Region for Parent
Private Sub AssignNewRegion(ByVal hwndParent As Long)

    SetWindowRgn hwndParent, hRgn(), True

End Sub

