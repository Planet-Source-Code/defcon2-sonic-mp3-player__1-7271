VERSION 5.00
Begin VB.Form frmEditor 
   BorderStyle     =   3  'Fester Dialog
   ClientHeight    =   4770
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4065
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   4065
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton cmdBack 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   20
      TabIndex        =   27
      Top             =   360
      Width           =   255
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3780
      TabIndex        =   26
      Top             =   360
      Width           =   255
   End
   Begin VB.Frame FrameEditor 
      Height          =   400
      Left            =   20
      TabIndex        =   23
      Top             =   -80
      Width           =   4040
      Begin VB.CommandButton cmdClose 
         Caption         =   "Î"
         BeginProperty Font 
            Name            =   "Wingdings 2"
            Size            =   9
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3720
         TabIndex        =   25
         Top             =   145
         Width           =   255
      End
      Begin VB.Label lblEditorCaption 
         Caption         =   "Eigenschaften von:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   135
         Width           =   3615
      End
   End
   Begin VB.TextBox txtEditMode 
      Alignment       =   2  'Zentriert
      Appearance      =   0  '2D
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox txtEditFreq 
      Alignment       =   2  'Zentriert
      Appearance      =   0  '2D
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   885
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox txtEditLayer 
      Alignment       =   2  'Zentriert
      Appearance      =   0  '2D
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox txtEditBitrate 
      Alignment       =   2  'Zentriert
      Appearance      =   0  '2D
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   885
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Abbrechen"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3000
      TabIndex        =   14
      Top             =   4320
      Width           =   975
   End
   Begin VB.ComboBox cmbGenre 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1125
      TabIndex        =   13
      Text            =   "Combo1"
      Top             =   3360
      Width           =   2895
   End
   Begin VB.TextBox txtEditComment 
      Appearance      =   0  '2D
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1125
      TabIndex        =   12
      Top             =   3840
      Width           =   2895
   End
   Begin VB.TextBox txtEditYear 
      Appearance      =   0  '2D
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1125
      MaxLength       =   4
      TabIndex        =   8
      Top             =   3000
      Width           =   735
   End
   Begin VB.TextBox txtEditAlbum 
      Appearance      =   0  '2D
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1125
      TabIndex        =   6
      Top             =   2640
      Width           =   2895
   End
   Begin VB.TextBox txtEditorFileName 
      Appearance      =   0  '2D
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   285
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "FileName"
      Top             =   360
      Width           =   3450
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Speichern"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1920
      TabIndex        =   4
      Top             =   4320
      Width           =   975
   End
   Begin VB.TextBox txtEditSongname 
      Appearance      =   0  '2D
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1125
      TabIndex        =   1
      Top             =   2280
      Width           =   2895
   End
   Begin VB.TextBox txtEditArtist 
      Appearance      =   0  '2D
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1125
      TabIndex        =   0
      Top             =   1920
      Width           =   2895
   End
   Begin VB.Label lblEditMode 
      Caption         =   " Mode:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   22
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label lblEditFreq 
      Caption         =   " Frequenz:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   45
      TabIndex        =   20
      Top             =   1200
      Width           =   855
   End
   Begin VB.Line Line1 
      X1              =   45
      X2              =   4020
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label lblEditLayer 
      Caption         =   " MPEG:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   18
      Top             =   840
      Width           =   855
   End
   Begin VB.Label lblEditBitrate 
      Caption         =   " Bitrate:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   45
      TabIndex        =   16
      Top             =   840
      Width           =   855
   End
   Begin VB.Label lblEditComment 
      Caption         =   " Kommentar:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   45
      TabIndex        =   11
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label lblEditGenre 
      Caption         =   " Genre:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   45
      TabIndex        =   10
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label lblEditYear 
      Caption         =   " Jahr:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   45
      TabIndex        =   9
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label lblEditAlbum 
      Caption         =   " Album:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   45
      TabIndex        =   7
      Top             =   2640
      Width           =   975
   End
   Begin VB.Line Line 
      X1              =   45
      X2              =   4020
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label lblEditSongname 
      Caption         =   " Titel:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   45
      TabIndex        =   3
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label lblEditArtist 
      Caption         =   " Artist:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   45
      TabIndex        =   2
      Top             =   1920
      Width           =   975
   End
End
Attribute VB_Name = "frmEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdBack_Click()
        
    WriteTag txtEditorFileName.Text, txtEditSongname.Text, txtEditArtist.Text, txtEditAlbum.Text, txtEditYear.Text, txtEditComment.Text, cmbGenre.Text
    ReadMP3 txtEditorFileName.Text, True, True
    Artist = GetMP3Info.Artist
    Songname = GetMP3Info.Songname
    frmPlayer.lstFiles.List(frmPlayer.lstFiles.ListIndex) = frmPlayer.lstFiles.ListIndex + 1 & ". " & Artist & " - " & Songname
        
        
    If frmPlayer.lstFiles.Selected(0) = True Then Exit Sub
    
    frmPlayer.lstFiles.Selected(frmPlayer.lstFiles.ListIndex - 1) = True
    
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

Private Sub cmdCancel_Click()
        
    frmPlayer.timerHide.Enabled = True
    Unload frmEditor
    
End Sub

Private Sub cmdClose_Click()
    
    frmPlayer.timerHide.Enabled = True
    Unload frmEditor
    
End Sub

Private Sub cmdNext_Click()
        
    WriteTag txtEditorFileName.Text, txtEditSongname.Text, txtEditArtist.Text, txtEditAlbum.Text, txtEditYear.Text, txtEditComment.Text, cmbGenre.Text
    ReadMP3 txtEditorFileName.Text, True, True
    Artist = GetMP3Info.Artist
    Songname = GetMP3Info.Songname
    frmPlayer.lstFiles.List(frmPlayer.lstFiles.ListIndex) = frmPlayer.lstFiles.ListIndex + 1 & ". " & Artist & " - " & Songname
        
        
    If frmPlayer.lstFiles.Selected(frmPlayer.lstFiles.ListCount - 1) = True Then Exit Sub
    
    frmPlayer.lstFiles.Selected(frmPlayer.lstFiles.ListIndex + 1) = True
    
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

Private Sub cmdOK_Click()
    
    ' ## Ersten Buchstaben entfernen
    'txtEditSongname.Text = Mid(txtEditSongname.Text, 2) & Left(txtEditSongname.Text, 1)
    ' ## Ende
    
    WriteTag txtEditorFileName.Text, txtEditSongname.Text, txtEditArtist.Text, txtEditAlbum.Text, txtEditYear.Text, txtEditComment.Text, cmbGenre.Text
    ReadMP3 txtEditorFileName.Text, True, True
    Artist = GetMP3Info.Artist
    Songname = GetMP3Info.Songname
    frmPlayer.lstFiles.List(frmPlayer.lstFiles.ListIndex) = frmPlayer.lstFiles.ListIndex + 1 & ". " & Artist & " - " & Songname
    
    frmPlayer.timerHide.Enabled = True
    Unload frmEditor
    
End Sub


Private Sub Form_Load()
    
    frmEditor.Top = frmEinstellungen.txtMenuY.Text
    frmEditor.Left = frmEinstellungen.txtMenuX.Text
    AlwaysOnTop frmEditor, 1
    
    If frmEinstellungen.chkWin2000Effects.Value = 1 Then
    
    'Windows 2000 Effects
    
    Dim ps As Integer

    If Not EnableTransparanty(Me.hWnd, 0) = 0 Then
        'MsgBox "Error using Transparety:" & Chr(10) & "Kein Windows 2000 erkannt!" & Chr(10) & Chr(10) & "Showing form without an effect!", vbCritical
        'frmEinstellungen.Show
        Me.Show
    Else
        Me.Enabled = False
        Me.Show
        DoEvents
    
    For ps = 0 To 245 Step 13
        DoEvents
        Call EnableTransparanty(Me.hWnd, ps)
        DoEvents
    Next
        Me.Enabled = True
        
    End If
    End If
    
End Sub



Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        
    txtEditArtist.BorderStyle = 0
    lblEditArtist.BackColor = &H8000000F
    lblEditArtist.ForeColor = &H80000012
    txtEditSongname.BorderStyle = 0
    lblEditSongname.BackColor = &H8000000F
    lblEditSongname.ForeColor = &H80000012
    txtEditAlbum.BorderStyle = 0
    lblEditAlbum.BackColor = &H8000000F
    lblEditAlbum.ForeColor = &H80000012
    txtEditYear.BorderStyle = 0
    lblEditYear.BackColor = &H8000000F
    lblEditYear.ForeColor = &H80000012
    txtEditComment.BorderStyle = 0
    lblEditComment.BackColor = &H8000000F
    lblEditComment.ForeColor = &H80000012
    
End Sub





Private Sub lblEditArtist_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    txtEditArtist.SetFocus
    txtEditArtist.BorderStyle = 1
    lblEditArtist.BackColor = &H8000000D
    lblEditArtist.ForeColor = &HFFFFFF
    ' ## OFF
    txtEditSongname.BorderStyle = 0
    lblEditSongname.BackColor = &H8000000F
    lblEditSongname.ForeColor = &H80000012
    txtEditAlbum.BorderStyle = 0
    lblEditAlbum.BackColor = &H8000000F
    lblEditAlbum.ForeColor = &H80000012
    txtEditYear.BorderStyle = 0
    lblEditYear.BackColor = &H8000000F
    lblEditYear.ForeColor = &H80000012
    txtEditComment.BorderStyle = 0
    lblEditComment.BackColor = &H8000000F
    lblEditComment.ForeColor = &H80000012
    
End Sub



Private Sub lblEditorCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    ReleaseCapture
    SendMessage frmEditor.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
    
End Sub

Private Sub txtEditArtist_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    txtEditArtist.SetFocus
    txtEditArtist.BorderStyle = 1
    lblEditArtist.BackColor = &H8000000D
    lblEditArtist.ForeColor = &HFFFFFF
    ' ## OFF
    txtEditSongname.BorderStyle = 0
    lblEditSongname.BackColor = &H8000000F
    lblEditSongname.ForeColor = &H80000012
    txtEditAlbum.BorderStyle = 0
    lblEditAlbum.BackColor = &H8000000F
    lblEditAlbum.ForeColor = &H80000012
    txtEditYear.BorderStyle = 0
    lblEditYear.BackColor = &H8000000F
    lblEditYear.ForeColor = &H80000012
    txtEditComment.BorderStyle = 0
    lblEditComment.BackColor = &H8000000F
    lblEditComment.ForeColor = &H80000012
    
End Sub

Private Sub lblEditSongname_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    txtEditSongname.SetFocus
    txtEditSongname.BorderStyle = 1
    lblEditSongname.BackColor = &H8000000D
    lblEditSongname.ForeColor = &HFFFFFF
    ' ## OFF
    txtEditArtist.BorderStyle = 0
    lblEditArtist.BackColor = &H8000000F
    lblEditArtist.ForeColor = &H80000012
    txtEditAlbum.BorderStyle = 0
    lblEditAlbum.BackColor = &H8000000F
    lblEditAlbum.ForeColor = &H80000012
    txtEditYear.BorderStyle = 0
    lblEditYear.BackColor = &H8000000F
    lblEditYear.ForeColor = &H80000012
    txtEditComment.BorderStyle = 0
    lblEditComment.BackColor = &H8000000F
    lblEditComment.ForeColor = &H80000012
    
End Sub


Private Sub txtEditSongname_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    txtEditSongname.SetFocus
    txtEditSongname.BorderStyle = 1
    lblEditSongname.BackColor = &H8000000D
    lblEditSongname.ForeColor = &HFFFFFF
    ' ## OFF
    txtEditArtist.BorderStyle = 0
    lblEditArtist.BackColor = &H8000000F
    lblEditArtist.ForeColor = &H80000012
    txtEditAlbum.BorderStyle = 0
    lblEditAlbum.BackColor = &H8000000F
    lblEditAlbum.ForeColor = &H80000012
    txtEditYear.BorderStyle = 0
    lblEditYear.BackColor = &H8000000F
    lblEditYear.ForeColor = &H80000012
    txtEditComment.BorderStyle = 0
    lblEditComment.BackColor = &H8000000F
    lblEditComment.ForeColor = &H80000012
    
End Sub

Private Sub lblEditAlbum_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    txtEditAlbum.SetFocus
    txtEditAlbum.BorderStyle = 1
    lblEditAlbum.BackColor = &H8000000D
    lblEditAlbum.ForeColor = &HFFFFFF
    ' ## OFF
    txtEditArtist.BorderStyle = 0
    lblEditArtist.BackColor = &H8000000F
    lblEditArtist.ForeColor = &H80000012
    txtEditSongname.BorderStyle = 0
    lblEditSongname.BackColor = &H8000000F
    lblEditSongname.ForeColor = &H80000012
    txtEditYear.BorderStyle = 0
    lblEditYear.BackColor = &H8000000F
    lblEditYear.ForeColor = &H80000012
    txtEditComment.BorderStyle = 0
    lblEditComment.BackColor = &H8000000F
    lblEditComment.ForeColor = &H80000012
    
End Sub


Private Sub txtEditAlbum_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        
    txtEditAlbum.SetFocus
    txtEditAlbum.BorderStyle = 1
    lblEditAlbum.BackColor = &H8000000D
    lblEditAlbum.ForeColor = &HFFFFFF
    ' ## OFF
    txtEditArtist.BorderStyle = 0
    lblEditArtist.BackColor = &H8000000F
    lblEditArtist.ForeColor = &H80000012
    txtEditSongname.BorderStyle = 0
    lblEditSongname.BackColor = &H8000000F
    lblEditSongname.ForeColor = &H80000012
    txtEditYear.BorderStyle = 0
    lblEditYear.BackColor = &H8000000F
    lblEditYear.ForeColor = &H80000012
    txtEditComment.BorderStyle = 0
    lblEditComment.BackColor = &H8000000F
    lblEditComment.ForeColor = &H80000012
    
End Sub

Private Sub lblEditYear_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    txtEditYear.SetFocus
    txtEditYear.BorderStyle = 1
    lblEditYear.BackColor = &H8000000D
    lblEditYear.ForeColor = &HFFFFFF
    ' ## OFF
    txtEditArtist.BorderStyle = 0
    lblEditArtist.BackColor = &H8000000F
    lblEditArtist.ForeColor = &H80000012
    txtEditSongname.BorderStyle = 0
    lblEditSongname.BackColor = &H8000000F
    lblEditSongname.ForeColor = &H80000012
    txtEditAlbum.BorderStyle = 0
    lblEditAlbum.BackColor = &H8000000F
    lblEditAlbum.ForeColor = &H80000012
    txtEditComment.BorderStyle = 0
    lblEditComment.BackColor = &H8000000F
    lblEditComment.ForeColor = &H80000012
    
End Sub


Private Sub txtEditYear_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    txtEditYear.SetFocus
    txtEditYear.BorderStyle = 1
    lblEditYear.BackColor = &H8000000D
    lblEditYear.ForeColor = &HFFFFFF
    ' ## OFF
    txtEditArtist.BorderStyle = 0
    lblEditArtist.BackColor = &H8000000F
    lblEditArtist.ForeColor = &H80000012
    txtEditSongname.BorderStyle = 0
    lblEditSongname.BackColor = &H8000000F
    lblEditSongname.ForeColor = &H80000012
    txtEditAlbum.BorderStyle = 0
    lblEditAlbum.BackColor = &H8000000F
    lblEditAlbum.ForeColor = &H80000012
    txtEditComment.BorderStyle = 0
    lblEditComment.BackColor = &H8000000F
    lblEditComment.ForeColor = &H80000012
    
End Sub

Private Sub lblEditComment_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    txtEditComment.SetFocus
    txtEditComment.BorderStyle = 1
    lblEditComment.BackColor = &H8000000D
    lblEditComment.ForeColor = &HFFFFFF
    ' ## OFF
    txtEditArtist.BorderStyle = 0
    lblEditArtist.BackColor = &H8000000F
    lblEditArtist.ForeColor = &H80000012
    txtEditSongname.BorderStyle = 0
    lblEditSongname.BackColor = &H8000000F
    lblEditSongname.ForeColor = &H80000012
    txtEditAlbum.BorderStyle = 0
    lblEditAlbum.BackColor = &H8000000F
    lblEditAlbum.ForeColor = &H80000012
    txtEditYear.BorderStyle = 0
    lblEditYear.BackColor = &H8000000F
    lblEditYear.ForeColor = &H80000012
    
End Sub


Private Sub txtEditComment_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        
    txtEditComment.SetFocus
    txtEditComment.BorderStyle = 1
    lblEditComment.BackColor = &H8000000D
    lblEditComment.ForeColor = &HFFFFFF
    ' ## OFF
    txtEditArtist.BorderStyle = 0
    lblEditArtist.BackColor = &H8000000F
    lblEditArtist.ForeColor = &H80000012
    txtEditSongname.BorderStyle = 0
    lblEditSongname.BackColor = &H8000000F
    lblEditSongname.ForeColor = &H80000012
    txtEditAlbum.BorderStyle = 0
    lblEditAlbum.BackColor = &H8000000F
    lblEditAlbum.ForeColor = &H80000012
    txtEditYear.BorderStyle = 0
    lblEditYear.BackColor = &H8000000F
    lblEditYear.ForeColor = &H80000012
    
End Sub

