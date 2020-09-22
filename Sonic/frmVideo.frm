VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Begin VB.Form frmVideo 
   Caption         =   "Form1"
   ClientHeight    =   4905
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7875
   LinkTopic       =   "Form1"
   ScaleHeight     =   4905
   ScaleWidth      =   7875
   StartUpPosition =   3  'Windows-Standard
   Visible         =   0   'False
   Begin VB.Timer timerPlayer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6000
      Top             =   240
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer 
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5655
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   0   'False
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   0   'False
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -950
      WindowlessVideo =   0   'False
   End
End
Attribute VB_Name = "frmVideo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    ReleaseCapture
    SendMessage frmVideo.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    frmEinstellungen.txtVideoPosX.Text = frmVideo.Top
    frmEinstellungen.txtVideoPosY.Text = frmVideo.Left
    
    frmEinstellungen.txtVideoH.Text = frmVideo.Height
    frmEinstellungen.txtVideoW.Text = frmVideo.Width
        
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    frmEinstellungen.chkVideo.Value = 0
        
End Sub

Private Sub Form_Resize()
    
    frmVideo.MediaPlayer.Width = frmVideo.Width - 440
    frmVideo.MediaPlayer.Height = frmVideo.Height - 640
        
End Sub


Private Sub picAnzeige_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    
    If Data.GetFormat(vbCFFiles) Then
        Effect = vbDropEffectCopy
    Else
        Effect = vbDropEffectNone
    End If
    
End Sub
    

Private Sub MediaPlayer_EndOfStream(ByVal Result As Long)
        
    If frmEinstellungen.chkRepeat.Value = 0 Then Exit Sub
            
    ' ## Überprüfe ob schon am Ende
    If frmPlayer.lstFiles.ListIndex = frmPlayer.lstFiles.ListCount - 1 Then frmPlayer.txtCurrentTitle.Text = "0"
    ' ## Wähle nächstes Item aus Liste aus oder shuffle
    
    If frmEinstellungen.chkShuffle.Value = 1 Then
    ShuffleFile = Int(((frmPlayer.lstFiles.ListCount - 1) * Rnd) + 1)
    frmPlayer.lstFiles.Selected(ShuffleFile) = True
    Else
    frmPlayer.lstFiles.Selected(frmPlayer.txtCurrentTitle.Text) = True
    End If
    
    If frmPlayer.lstFiles.List(frmPlayer.lstFiles.ListIndex) = "" Then
        ShowMsgBox "Kein Song gewählt:", "Bitte wählen Sie einen Song aus der abgespielt werden soll."
        
        Exit Sub
    End If
    
    Dim File, Tag As String
               
    
    File = frmPlayer.lstFileNames.List(frmPlayer.lstFileNames.ListIndex)
    
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
    'frmPlayer.lstFiles.SetFocus
    
    frmPlayer.Slider.Enabled = True
    ' ## Zähle in der Liste eins weiter
    frmPlayer.txtCurrentTitle.Text = frmPlayer.lstFiles.ListIndex + 1
    frmVideo.Caption = File
    frmVideo.MediaPlayer.Open File
    ' ## Icon in der Anzeige ändern
    frmPlayer.lblStatIcon.Caption = "4"
    frmPlayer.lblTitle.Caption = " " & frmPlayer.lstFiles.List(frmPlayer.lstFiles.ListIndex)
    
    Pause 0.5
    ' ## Slider auf Position null stellen
    frmPlayer.Slider.Min = frmVideo.MediaPlayer.SelectionStart
    'frmPlayer.lstFiles.SetFocus
    
End Sub



Private Sub MediaPlayer_PlayStateChange(ByVal OldState As Long, ByVal NewState As Long)
    
    ' ## Slider Max auf Position Ende stellen
    frmPlayer.Slider.Max = frmVideo.MediaPlayer.SelectionEnd
    
    Pos = frmPlayer.Slider.Max
    Min = Int(Pos / 60): Sec = Int(Pos - Min * 60)
    Zeit = Format$(Min, "00") + ":" + Format$(Sec, "00")
    
    'frmPlayer.txtTitle.Text = frmMain.lstFiles.List(frmMain.lstFiles.ListIndex) & "  [" & Zeit & "]  _-_  [sonic]  "
    'frmPlayer.Timer.Enabled = True
    
    
End Sub

Private Sub timerPlayer_Timer()
        
    frmPlayer.Slider.Value = frmVideo.MediaPlayer.CurrentPosition
    ' ## Time normal
    Pos = frmVideo.MediaPlayer.CurrentPosition
    Min = Int(Pos / 60): Sec = Int(Pos - Min * 60)
    ' ## Time rückwärts
    Pos2 = frmVideo.MediaPlayer.SelectionEnd - frmVideo.MediaPlayer.CurrentPosition
    Min2 = Int(Pos2 / 60): sec2 = Int(Pos2 - Min2 * 60)


    Track = frmPlayer.txtCurrentTitle.Text
    If Track < 10 Then Track = "0" & Track
    frmPlayer.lblTime.Caption = "[" & Track & "]  " & Format$(Min, "00") + ":" + Format$(Sec, "00")

    frmPlayer.lblTime2.Caption = "[" & Track & "] -" & Format$(Min2, "00") + ":" + Format$(sec2, "00")
    
End Sub

