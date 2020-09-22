VERSION 5.00
Begin VB.Form frmOwnMnu 
   Caption         =   "MenuIcons"
   ClientHeight    =   2715
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   ScaleHeight     =   2715
   ScaleWidth      =   6045
   StartUpPosition =   3  'Windows-Standard
   Begin VB.PictureBox img 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   9
      Left            =   120
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   8
      Top             =   1680
      Width           =   240
   End
   Begin VB.PictureBox tmpPicture 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'Kein
      Height          =   255
      Left            =   5760
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   7
      Top             =   120
      Width           =   255
      Begin VB.PictureBox img 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'Kein
         Height          =   240
         Index           =   0
         Left            =   0
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   9
         Top             =   0
         Width           =   240
         Begin VB.PictureBox img 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BorderStyle     =   0  'Kein
            Height          =   240
            Index           =   1
            Left            =   0
            ScaleHeight     =   16
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   16
            TabIndex        =   10
            Top             =   0
            Width           =   240
         End
      End
   End
   Begin VB.PictureBox img 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   8
      Left            =   120
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   6
      Top             =   1560
      Width           =   240
   End
   Begin VB.PictureBox img 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   7
      Left            =   120
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   5
      Top             =   1320
      Width           =   240
   End
   Begin VB.PictureBox img 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   6
      Left            =   120
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   4
      Top             =   1080
      Width           =   240
   End
   Begin VB.PictureBox img 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   5
      Left            =   120
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   3
      Top             =   840
      Width           =   240
   End
   Begin VB.PictureBox img 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   4
      Left            =   120
      Picture         =   "Form1.frx":0342
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   2
      Top             =   600
      Width           =   240
   End
   Begin VB.PictureBox img 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   3
      Left            =   120
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   1
      Top             =   360
      Width           =   240
   End
   Begin VB.PictureBox img 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   2
      Left            =   120
      Picture         =   "Form1.frx":0684
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   0
      Top             =   120
      Width           =   240
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&Menu0"
      Begin VB.Menu mnuPlay 
         Caption         =   "&Abspielen"
      End
      Begin VB.Menu mnuSeperator0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRemove 
         Caption         =   "&Entfernen"
      End
      Begin VB.Menu mnuRemoveAll 
         Caption         =   "&Alle Entfernen"
      End
      Begin VB.Menu mnuSeperator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddtoBookmark 
         Caption         =   "&Als Lesezeichen"
      End
      Begin VB.Menu sep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditor 
         Caption         =   "&TAG bearbeiten"
      End
   End
End
Attribute VB_Name = "frmOwnMnu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit

Dim pnt As PaintEffects

Dim MyFont As Long
Dim OldFont As Long

Dim wlOldProc As Long

Dim Caps(2 To 27) As String

Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
Private Sub Form_Load()

     Set pnt = New PaintEffects

     Caps(2) = "Abspielen"
     Caps(3) = ""
     Caps(4) = "Entfernen"
     Caps(5) = "Alle Entfernen"
     Caps(6) = ""
     Caps(7) = "Als Lesezeichen"
     Caps(8) = ""
     Caps(9) = "TAG bearbeiten"
         
     enableIcons

End Sub



Public Function MsgProc(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'This procedure is called because we've subclassed
'this form. We will catch DRAWITEM and MEASUREITEM
'messages and pass the rest of them on.

'Various structs we'll need
Dim MeasureInfo As MEASUREITEMSTRUCT
Dim DrawInfo As DRAWITEMSTRUCT
Dim mii As MENUITEMINFO
'Set later for separator flag:
Dim IsSep As Boolean
'Our custom brush and the old one
Dim hBr As Long, hOldBr As Long
'Our custom pen and the old one
Dim hPen As Long, hOldPen As Long
'The text color of the menu items
Dim lTextColor As Long
'Now much to bump the menu's selection
'rectangle over
Dim iRectOffset As Integer
Dim isChecked As Boolean

If wMsg = WM_DRAWITEM Then
     If wParam = 0 Then 'It was sent by the menu
          'Get DRAWINFOSTRUCT -- copy it to our
          'empty structure from the pointer in lParam
          Call CopyMem(DrawInfo, ByVal lParam, LenB(DrawInfo))
          IsSep = IsSeparator(DrawInfo.itemID)

          '===Set the menu font through its hDC...===
          MyFont = SendMessage(Me.hWnd, WM_GETFONT, 0&, 0&)
          OldFont = SelectObject(DrawInfo.hdc, MyFont)
          'We draw the item based on Un/Selected:
        
          'Some constants can be interpeted as others
          
          Select Case DrawInfo.itemState
          
                Case 257, 289
                    DrawInfo.itemState = ODS_SELECTED

                Case 9, 297
                    DrawInfo.itemState = 265
                    
                Case 294, 295, 7, 6
                    DrawInfo.itemState = 262
                    
                Case 303, 302
                    DrawInfo.itemState = 270
                    
                Case 296
                    DrawInfo.itemState = 264
                    
                Case 14, 15
                    DrawInfo.itemState = 271

          End Select
          
          If DrawInfo.itemState = ODS_SELECTED Or DrawInfo.itemState = 265 Then
               hBr = CreateSolidBrush(GetSysColor(COLOR_HIGHLIGHT))
               hPen = GetPen(1, GetSysColor(COLOR_HIGHLIGHT))
               lTextColor = GetSysColor(COLOR_HIGHLIGHTTEXT)
          Else
               hBr = CreateSolidBrush(GetSysColor(COLOR_MENU))
               hPen = GetPen(1, GetSysColor(COLOR_MENU))
               lTextColor = GetSysColor(COLOR_MENUTEXT)
          End If
          
          'We're going to draw on the menu
          QuickGDI.TargethDC = DrawInfo.hdc
          'Select our new, correctly colored objects:
          
          hOldBr = SelectObject(DrawInfo.hdc, hBr)
          hOldPen = SelectObject(DrawInfo.hdc, hPen)
          
          With DrawInfo.rcItem
            If DrawInfo.itemState <> ODS_SELECTED And DrawInfo.itemState <> 265 Then
                        'Clear the space where the image is
                   QuickGDI.DrawRect .Left, .Top, 22, .Bottom
            End If
          
            'Check to see if the menu item is one of the
            'ones with a picture. If so, then we need to
            'move the edge of the drawing rectangle a little
            'to the left to make room for the image.
            iRectOffset = IIf(img(DrawInfo.itemID).picture <> 0, 23, 0)
            'Do we have a separator bar?
          
            If Not IsSep Then
                'Draw the rectangle onto the item's space
                QuickGDI.DrawRect .Left + iRectOffset, .Top, .Right, .Bottom
                'Print the item's text
                '(held in the Caps() array)
                Select Case DrawInfo.itemState
                
                Case 262, 263, 271, 270 'Disabled
                    lTextColor = GetSysColor(COLOR_WINDOW)
                    hPrint .Left + 25, .Top + 3, " " & Caps(DrawInfo.itemID), lTextColor
                    lTextColor = GetSysColor(COLOR_GRAYTEXT)
                    hPrint .Left + 24, .Top + 2, " " & Caps(DrawInfo.itemID), lTextColor
                    
                Case Else ' Object is enabled
                    hPrint .Left + 24, .Top + 2, " " & Caps(DrawInfo.itemID), lTextColor
                                        
                End Select
            End If
          End With
          
          'Select the old objects into the menu's DC
          Call SelectObject(DrawInfo.hdc, hOldBr)
          Call SelectObject(DrawInfo.hdc, hOldPen)
          
          'Delete the ones we created
          Call DeleteObject(hBr)
          Call DeleteObject(hPen)
          
          
          With DrawInfo
            'If the item had an image:
            If img(.itemID).picture.Handle <> 0 Then
            
               'If this item is selected, draw a raised
               'box around the image
               
               Select Case DrawInfo.itemState
               
               'HIER!
               
               Case 262, 263
               
                     Call buildEmbosedImage(img(.itemID))
                    
               Case 271, 270
               
                     ThreedBox 1, .rcItem.Top, 21, .rcItem.Bottom - 1, True
                     Call buildEmbosedImage(img(.itemID))
                                                  
               Case 265, 264, 8
               
                     ThreedBox 1, .rcItem.Top, 21, .rcItem.Bottom - 1, True
                     Call removeEmbosedImage(img(.itemID))
                    
               Case ODS_SELECTED
               
                     ThreedBox 1, .rcItem.Top, 21, .rcItem.Bottom - 1, False
                     Call removeEmbosedImage(img(.itemID))
                     
               Case Else
               
                    Call removeEmbosedImage(img(.itemID))
                    
               End Select
               
               If Not img(.itemID).Tag = "Embosed" Then
                pnt.PaintTransparentStdPic .hdc, 4, .rcItem.Top + 2, 16, 16, img(.itemID).picture, 0, 0, &HC0C0C0
               Else
                tmpPicture.picture = img(.itemID).Image
                pnt.PaintTransparentStdPic .hdc, 4, .rcItem.Top + 2, 16, 16, tmpPicture.picture, 0, 0, &HC0C0C0
               End If
                
          End If
            
          If IsSep Then
               'Draw the special separator bar
               ThreedBox .rcItem.Left, .rcItem.Top + 2, .rcItem.Right - 1, .rcItem.Bottom - 2, True
          End If
          End With
     End If
     'Don't pass this message on:
     MsgProc = False
     
     Exit Function
     

ElseIf wMsg = WM_MEASUREITEM Then
     'Get the MEASUREITEM struct from the pointer
     Call CopyMem(MeasureInfo, ByVal lParam, Len(MeasureInfo))
     IsSep = IsSeparator(MeasureInfo.itemID)
     'Tell Windows how big our items are.
     MeasureInfo.itemWidth = 140
     'If the item being measured is the separator
     'bar, the height should be 5 pixels, 18 if
     'otherwise...
     MeasureInfo.itemHeight = IIf(IsSep, 5, 20)
     'Return the information back to Windows
     Call CopyMem(ByVal lParam, MeasureInfo, Len(MeasureInfo))
     'Don't pass this message on:
     MsgProc = False
     Exit Function

End If

'We didn't handle this message,
'pass it on to the next WndProc
MsgProc = CallWindowProc(wlOldProc, hWnd, wMsg, wParam, lParam)

End Function


Private Sub enableIcons()
     If wlOldProc <> 0 Then Exit Sub

     Dim i As Integer

     MenuItems.MenuForm = Me
     
     'Start with File menu
     MenuItems.SubMenu = 0
     
     For i = 0 To 7
          MenuItems.MenuID = i
          OwnerDrawMenu (i + 2)
     Next
     
     'Next comes 2nd menu...
     MenuItems.SubMenu = 1
     For i = 0 To 4
          MenuItems.MenuID = i
           OwnerDrawMenu (i + 2)
     Next
     
     'Next comes 3th menu...
     MenuItems.SubMenu = 2
     For i = 0 To 4
          MenuItems.MenuID = i
          OwnerDrawMenu (i + 2)
     Next
     
     'Next comes 4th menu...
     MenuItems.SubMenu = 3
     For i = 0 To 3
          MenuItems.MenuID = i
          OwnerDrawMenu (i + 2)
     Next
     
     'Next comes 5th menu...
     MenuItems.SubMenu = 4
     For i = 0 To 3
          MenuItems.MenuID = i
          OwnerDrawMenu (i + 2)
     Next
 
     wlOldProc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf OwnMenuProc)

End Sub

Private Sub Form_Unload(Cancel As Integer)
     
     If wlOldProc <> 0 Then
          SetWindowLong hWnd, GWL_WNDPROC, wlOldProc
     End If
     Set pnt = Nothing

     'Destroy the font object created in
     'the form's window procedure.
     Call DeleteObject(MyFont)

End Sub

Private Sub mnuEditor_Click()
                
        Ext = GetExtension(frmPlayer.lstFileNames.List(frmPlayer.lstFileNames.ListIndex))
        
        Select Case Ext
            Case "mp3", "Mp3", "MP3"
            
        frmEditor.Show
                
        ' ## Tag einlesen
        Call ReadMP3(frmPlayer.lstFileNames.List(frmPlayer.lstFileNames.ListIndex), True, True)
        Artist = GetMP3Info.Artist
        Songname = GetMP3Info.Songname
        Album = GetMP3Info.Album
        Jahr = GetMP3Info.Year
        Genre = GetMP3Info.Genre
        Comment = GetMP3Info.Comment
        Bitrate = GetMP3Info.Bitrate & "kBit"
        MpegLayer = GetMP3Info.MpegVersion & ".0 Layer " & GetMP3Info.MpegLayer
        Freq = GetMP3Info.Frequency & "Hz"
        Mode = GetMP3Info.Mode
        
        'MsgBox GetMP3Info.CopyRight
        'MsgBox GetMP3Info.CRC
        'MsgBox GetMP3Info.Duration  'Länge
        
        frmEditor.txtEditorFileName.Text = frmPlayer.lstFileNames.List(frmPlayer.lstFileNames.ListIndex)
        frmEditor.txtEditArtist.Text = Artist
        frmEditor.txtEditSongname.Text = Songname
        frmEditor.txtEditAlbum.Text = Album
        frmEditor.txtEditYear.Text = Jahr
        frmEditor.cmbGenre.Text = Genre
        frmEditor.txtEditComment.Text = Comment
        
        frmEditor.txtEditBitrate.Text = Bitrate
        frmEditor.txtEditLayer.Text = MpegLayer
        frmEditor.txtEditFreq.Text = Freq
        frmEditor.txtEditMode.Text = Mode
        
        frmPlayer.timerHide.Enabled = False
        
        Case Else
            ShowMsgBox "Keine MP3 Datei:", "Die gewählte Datei ist keine MP3 und hat auch somit keinen ID3 Tag."
        End Select
End Sub

Private Sub mnuPlay_Click()
    
    Dim File, Tag As String
    File = frmPlayer.lstFileNames.List(frmPlayer.lstFileNames.ListIndex)
        
    frmPlayer.Slider.Enabled = True
    frmPlayer.txtCurrentTitle.Text = frmPlayer.lstFiles.ListIndex + 1
    
    frmVideo.MediaPlayer.Open File
    frmVideo.timerPlayer.Enabled = True
    frmPlayer.Slider.Min = frmVideo.MediaPlayer.SelectionStart
    
    ' ## Anzeige
    frmPlayer.lblStatIcon.Caption = "4"
    frmPlayer.lblTitle.Caption = " " & frmPlayer.lstFiles.List(frmPlayer.lstFiles.ListIndex)
    
    
    ' ## Extension
    Ext = GetExtension(frmPlayer.lstFileNames.List(frmPlayer.lstFileNames.ListIndex))
    
    Select Case Ext
        Case "mp3", "Mp3", "MP3"
            frmVideo.Hide
        Case "mpg", "Mpg", "MPG"
            frmVideo.Show
        Case "avi", "Avi", "AVI"
            frmVideo.Show
        Case "mpeg", "Mpeg", "MPEG"
            frmVideo.Show
        Case "asf", "Asf", "ASF"
            frmVideo.Show
    End Select
    
    frmPlayer.SliderVolume.Value = frmEinstellungen.txtVol.Text
    frmPlayer.lstFiles.SetFocus
    
End Sub

Private Sub mnuRemove_Click()
    
    SEL = frmPlayer.lstFiles.ListIndex
    frmPlayer.lstFiles.RemoveItem SEL
    frmPlayer.lstFileNames.RemoveItem SEL
    
End Sub

Private Sub mnuRemoveAll_Click()
    
    frmPlayer.lstFiles.Clear
    frmPlayer.lstFileNames.Clear
    
End Sub
