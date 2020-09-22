Attribute VB_Name = "MenuItems"
Option Explicit
Dim hMenu As Long
Dim hSubMenu As Long
Dim mnuID As Long

Dim m_Form As frmOwnMnu

'Subclassing stuff we'll need...
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Const GWL_WNDPROC = (-4)
'Messages to use in the wndproc
Public Const WM_DRAWITEM = &H2B
Public Const WM_MEASUREITEM = &H2C
Public Const WM_MENUSELECT = &H11F
Public Const WM_COMMAND = &H111
Public Const WM_GETFONT = &H31

Type MENUITEMINFO
     cbSize As Long
     fMask As Long
     fType As Long
     fState As Long
     wID As Long
     hSubMenu As Long
     hbmpChecked As Long
     hbmpUnchecked As Long
     dwItemData As Long
     dwTypeData As Long
     cch As Long
End Type
Public Const MIIM_TYPE = &H10

Type MEASUREITEMSTRUCT
     CtlType As Long
     CtlID As Long
     itemID As Long
     itemWidth As Long
     itemHeight As Long
     ItemData As Long
End Type
Type DRAWITEMSTRUCT
     CtlType As Long
     CtlID As Long
     itemID As Long
     itemAction As Long
     itemState As Long
     hwndItem As Long
     hdc As Long
     rcItem As RECT
     ItemData As Long
End Type

Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long

Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Long) As Long

Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal uItem As Long, ByVal ByPosition As Long, lpMenuItemInfo As MENUITEMINFO) As Boolean

Private Const MF_BYCOMMAND = &H0
Private Const MF_BYPOSITION = &H400
Private Const MF_OWNERDRAW = &H100
Private Const MF_SEPARATOR = &H800
Public Const MFT_SEPARATOR = MF_SEPARATOR

Public Const ODS_SELECTED = &H1


Public Property Get MenuForm() As frmOwnMnu
     Set MenuForm = m_Form
End Property
Public Property Let MenuForm(ByVal vNewValue As frmOwnMnu)
     Set m_Form = vNewValue
     hMenu = GetMenu(m_Form.hWnd)
End Property

Public Property Get MenuID() As Long
     MenuID = mnuID
End Property
Public Property Let MenuID(ByVal vNewValue As Long)
     mnuID = GetMenuItemID(hSubMenu, vNewValue)
End Property

Public Sub OwnerDrawMenu(ByVal ItemData As Long)
     'Change the menu's style to owner-draw. You must
     'now subclass the form that this menu is on so
     'you can respond to the WM_MEASUREITEM and WM_DRAWITEM
     'messages.
     Dim mii As MENUITEMINFO
     mii.cbSize = Len(mii)
     mii.fMask = MIIM_TYPE
     GetMenuItemInfo hSubMenu, MenuID, False, mii
     If ((mii.fType And MF_SEPARATOR) = MF_SEPARATOR) Then
          '*Preserve* separator style...
          Call ModifyMenu(hSubMenu, MenuID, MF_BYCOMMAND Or MF_OWNERDRAW Or MF_SEPARATOR, MenuID, ItemData)
     Else
          Call ModifyMenu(hSubMenu, MenuID, MF_BYCOMMAND Or MF_OWNERDRAW, MenuID, ItemData)
     End If
End Sub

Public Function OwnMenuProc(ByVal hWnd As Long, _
ByVal wMsg As Long, ByVal wParam As Long, _
ByVal lParam As Long) As Long
     OwnMenuProc = frmOwnMnu.MsgProc(hWnd, wMsg, wParam, lParam)
End Function

Public Sub SetTopMenu(NewMnu As Long)
     hMenu = NewMnu
End Sub

Public Property Get SubMenu() As Long
     SubMenu = hSubMenu
End Property
Public Property Let SubMenu(ByVal vNewValue As Long)
     hSubMenu = GetSubMenu(hMenu, vNewValue)
End Property

Public Function buildEmbosedImage(picture As PictureBox)

If picture.Tag = "Embosed" Then Exit Function
Dim x As Integer, y As Integer
'
' Convert colors to grayscale
'
For x = 0 To picture.ScaleWidth
    For y = 0 To picture.ScaleHeight
    
        Select Case picture.Point(x, y)

        Case 16777215, &HC0C0C0
    
            picture.PSet (x, y), &HC0C0C0
            
        Case 0 To &H808080, 0
            
            picture.PSet (x, y), &H808080
            
        Case &H808080 To &HFF0000
        
            picture.PSet (x, y), &HC0C0C1
        
        Case &HFF0000 To &HFFFFFF
        
            picture.PSet (x, y), &HE0E0E0
                                
        End Select
    Next
Next
'
' Draw the white border
'
For x = picture.ScaleWidth To 1 Step -1
    For y = picture.ScaleHeight To 1 Step -1
        If picture.Point(x, y) = &HC0C0C0 Then
            If Not picture.Point(x - 1, y - 1) = &HC0C0C0 Then
                picture.PSet (x, y), vbWhite
            End If
        End If
    Next
Next
picture.Tag = "Embosed"

End Function
Public Function removeEmbosedImage(picture As PictureBox)
If picture.Tag = "Embosed" Then
    picture.Tag = ""
    picture.Cls
End If
End Function
Public Function IsSeparator(ByVal IID As Integer) As Boolean
     Dim mii As MENUITEMINFO
     mii.cbSize = Len(mii)
     mii.fMask = MIIM_TYPE
     mii.wID = IID
     GetMenuItemInfo GetMenu(m_Form.hWnd), IID, False, mii
     IsSeparator = ((mii.fType And MFT_SEPARATOR) = MFT_SEPARATOR)
End Function
Public Function HiWord(LongIn As Long) As Integer
     HiWord = (LongIn And &HFFFF0000) \ &H10000
End Function
Public Function LoWord(LongIn As Long) As Integer
     If (LongIn And &HFFFF&) > &H7FFF Then
          LoWord = (LongIn And &HFFFF&) - &H10000
     Else
          LoWord = LongIn And &HFFFF&
     End If
End Function
