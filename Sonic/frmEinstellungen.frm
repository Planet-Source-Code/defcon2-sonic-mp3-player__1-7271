VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEinstellungen 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Einstellungen"
   ClientHeight    =   12600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13335
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12600
   ScaleWidth      =   13335
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.TextBox txtListCount 
      Appearance      =   0  '2D
      Height          =   285
      Left            =   10320
      TabIndex        =   38
      Top             =   2400
      Width           =   255
   End
   Begin VB.Frame FrameInfo 
      Caption         =   "Information"
      Height          =   3255
      Left            =   240
      TabIndex        =   35
      Top             =   9000
      Visible         =   0   'False
      Width           =   6015
      Begin VB.Label lblInforamtion 
         Caption         =   $"frmEinstellungen.frx":0000
         Height          =   1215
         Left            =   240
         TabIndex        =   37
         Top             =   480
         Width           =   5175
      End
      Begin VB.Label lblHttp 
         BackStyle       =   0  'Transparent
         Caption         =   "http://www.defcon2.de"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   3960
         TabIndex        =   36
         Top             =   2880
         Width           =   1815
      End
   End
   Begin VB.Frame FramePlayer 
      Caption         =   "Player ... "
      Height          =   3255
      Left            =   240
      TabIndex        =   13
      Top             =   5640
      Visible         =   0   'False
      Width           =   6015
      Begin VB.PictureBox picSysIcon 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2760
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   27
         Top             =   2640
         Width           =   255
      End
      Begin VB.CheckBox chkVideoOnTop 
         Appearance      =   0  '2D
         Caption         =   "Video Immer im Vordergrund"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1080
         Width           =   2535
      End
      Begin VB.CheckBox chkPin 
         Appearance      =   0  '2D
         Caption         =   "Kein Docken"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1800
         Width           =   2175
      End
      Begin VB.CheckBox chkOnTop 
         Appearance      =   0  '2D
         Caption         =   "Player Immer im Vordergrund"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   2535
      End
      Begin VB.CheckBox chkSaveLastList 
         Appearance      =   0  '2D
         Caption         =   "Letzte Playliste speichern"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1440
         Width           =   2175
      End
      Begin VB.CheckBox chkWin2000Effects 
         Appearance      =   0  '2D
         Caption         =   "Windows 2000 Effekte (nur unter Windows 2000 oder höher)"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   4815
      End
      Begin VB.TextBox txtHideTime 
         Alignment       =   2  'Zentriert
         Appearance      =   0  '2D
         Height          =   285
         Left            =   4140
         MaxLength       =   2
         TabIndex        =   14
         Text            =   "3"
         Top             =   2160
         Width           =   375
      End
      Begin MSComctlLib.Slider SliderSysIcon 
         Height          =   255
         Left            =   1200
         TabIndex        =   26
         Top             =   2640
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         Min             =   1
         Max             =   4
         SelStart        =   4
         Value           =   4
      End
      Begin VB.Label lblSystrayIcon 
         Caption         =   "Systray Icon"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   2640
         Width           =   1935
      End
      Begin VB.Label lblHideTime 
         Caption         =   "Automatisches andocken am linken Bildschirmrand nach            Sekunden."
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   2205
         Width           =   5415
      End
   End
   Begin VB.Frame FrameStartUp 
      Caption         =   "Beim Start anzeigen ... "
      Height          =   3255
      Left            =   240
      TabIndex        =   10
      Top             =   1800
      Width           =   6015
      Begin VB.CheckBox chkVideo 
         Appearance      =   0  '2D
         Caption         =   "Video Fenster"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   2175
      End
      Begin VB.CheckBox chkSplash 
         Appearance      =   0  '2D
         Caption         =   "Splash Screen"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   2175
      End
   End
   Begin MSComctlLib.TabStrip TabStripEinstellungen 
      Height          =   3855
      Left            =   120
      TabIndex        =   34
      Top             =   1320
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   6800
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Start"
            Object.Tag             =   "Start"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Player"
            Object.Tag             =   "Player"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Info"
            Object.Tag             =   "Info"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtLastBrowseFolder 
      Appearance      =   0  '2D
      Height          =   285
      Left            =   11160
      TabIndex        =   31
      Top             =   3960
      Width           =   1815
   End
   Begin VB.CheckBox chkShuffle 
      Appearance      =   0  '2D
      Caption         =   "Shuffle"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   10800
      TabIndex        =   30
      Top             =   5640
      Width           =   2295
   End
   Begin VB.TextBox txtVideoW 
      Appearance      =   0  '2D
      Height          =   285
      Left            =   11760
      TabIndex        =   29
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox txtVideoH 
      Appearance      =   0  '2D
      Height          =   285
      Left            =   11760
      TabIndex        =   28
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox txtSystrayIcon 
      Appearance      =   0  '2D
      Height          =   285
      Left            =   12120
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton cmdHideSystray 
      Caption         =   "Hide SystrayIcon"
      Height          =   315
      Left            =   11640
      TabIndex        =   23
      Top             =   5160
      Width           =   1455
   End
   Begin VB.CommandButton cmdShowSystray 
      Caption         =   "ShowSystrayIcon"
      Height          =   315
      Left            =   11640
      TabIndex        =   22
      Top             =   4800
      Width           =   1455
   End
   Begin MSComctlLib.ImageList ImageListSysIcon 
      Left            =   11040
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEinstellungen.frx":0090
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEinstellungen.frx":139C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEinstellungen.frx":26A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEinstellungen.frx":4E5C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CheckBox chkRepeat 
      Appearance      =   0  '2D
      Caption         =   "Wiederholen"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   10800
      TabIndex        =   21
      Top             =   5880
      Width           =   2295
   End
   Begin VB.CheckBox chkDefault 
      Appearance      =   0  '2D
      Caption         =   "Standart INI und Playlisten"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   10800
      TabIndex        =   9
      Top             =   6120
      Width           =   2295
   End
   Begin VB.TextBox txtVol 
      Appearance      =   0  '2D
      Height          =   285
      Left            =   12600
      TabIndex        =   8
      Top             =   2760
      Width           =   375
   End
   Begin VB.TextBox txtMenuY 
      Appearance      =   0  '2D
      Height          =   285
      Left            =   12480
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox txtMenuX 
      Appearance      =   0  '2D
      Height          =   285
      Left            =   12480
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox txtPlaylistSelected 
      Appearance      =   0  '2D
      Height          =   285
      Left            =   10680
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   2400
      Width           =   2295
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   315
      Left            =   5160
      TabIndex        =   4
      Top             =   5280
      Width           =   975
   End
   Begin VB.TextBox txtVideoPosY 
      Appearance      =   0  '2D
      Height          =   285
      Left            =   11760
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox txtVideoPosX 
      Appearance      =   0  '2D
      Height          =   285
      Left            =   11760
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txtPlayerPosY 
      Appearance      =   0  '2D
      Height          =   285
      Left            =   11760
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox txtPlayerPosX 
      Appearance      =   0  '2D
      Height          =   285
      Left            =   11760
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblToopTipCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "Einstellungen ..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   240
      TabIndex        =   33
      Top             =   165
      Width           =   6015
   End
   Begin VB.Label lblToolTip 
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "sonic"
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   240
      TabIndex        =   32
      Top             =   480
      Width           =   6015
   End
   Begin VB.Shape Shape 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Undurchsichtig
      Height          =   1095
      Left            =   120
      Top             =   120
      Width           =   6255
   End
End
Attribute VB_Name = "frmEinstellungen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Window messages that identify mouse action
Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_MBUTTONDOWN = &H207
Private Const WM_MBUTTONUP = &H208
Private Const WM_MBUTTONDBLCLK = &H209

Dim WithEvents SysIcon As SYSTRAY
Attribute SysIcon.VB_VarHelpID = -1


Private Sub chkOnTop_Click()
    
    Select Case chkOnTop.Value
        Case 0
            AlwaysOnTop frmPlayer, 0
            frmPlayer.lblOnTop.ForeColor = &H808080
        Case 1
            AlwaysOnTop frmPlayer, 1
            frmPlayer.lblOnTop.ForeColor = &HE0E0E0
    End Select
    
End Sub


Private Sub chkOnTop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    lblToolTip.Caption = "Aktiviert, und das Player Fenster ist immer im Vordergrund."
    
End Sub



Private Sub chkPin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    lblToolTip.Caption = "Aktiviert, und das Player Fenster dockt nicht mehr am linken Bildschirmrand an."
    
End Sub

Private Sub chkRepeat_Click()
    
    Select Case chkRepeat.Value
        Case 0
            frmPlayer.lblRepeat.ForeColor = &H808080
        Case 1
            frmPlayer.lblRepeat.ForeColor = &HE0E0E0
    End Select
    
End Sub



Private Sub chkSaveLastList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    lblToolTip.Caption = "Speichert beim beenden automatisch die letzten Änderungen in jeder Playliste ab."
    
End Sub

Private Sub chkShuffle_Click()
    
    Select Case chkShuffle
        Case 0
            frmPlayer.lblShuffle.ForeColor = &H808080
        Case 1
            frmPlayer.lblShuffle.ForeColor = &HE0E0E0
    End Select
    
End Sub


Private Sub chkSplash_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    lblToolTip.Caption = "Splash Screen beim Start anzeigen."
    
End Sub

Private Sub chkVideo_Click()
    
    Select Case chkVideo.Value
        Case 0
            frmVideo.Hide
            frmPlayer.lblVideo.ForeColor = &H808080
        Case 1
            frmVideo.Show
            frmPlayer.lblVideo.ForeColor = &HE0E0E0
    End Select
    
End Sub

Private Sub chkVideo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    lblToolTip.Caption = "Video Fenster beim Start anzeigen."
    
End Sub

Private Sub chkVideoOnTop_Click()
    
    Select Case chkOnTop.Value
        Case 0
            AlwaysOnTop frmVideo, 0
        Case 1
            AlwaysOnTop frmVideo, 1
    End Select
    
End Sub

Private Sub chkVideoOnTop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    lblToolTip.Caption = "Aktiviert, und das Video Fenster ist immer im Vordergrund."
    
End Sub

Private Sub chkWin2000Effects_Click()
    
    Select Case chkWin2000Effects.Value
        Case 1
        
            Dim ps As Integer

            If Not EnableTransparanty(frmPlayer.hWnd, 0) = 0 Then
                'MsgBox "Error using Transparety:" & Chr(10) & "Kein Windows 2000 erkannt!" & Chr(10) & Chr(10) & "Showing form without an effect!", vbCritical
                'frmEinstellungen.Show
                'frmPlayer.Show
            Else
                frmPlayer.Enabled = False
                'frmPlayer.Show
                DoEvents
    
            For ps = 255 To 200 Step -3
                DoEvents
                Call EnableTransparanty(frmPlayer.hWnd, ps)
                DoEvents
            Next
                frmPlayer.Enabled = True
        
            End If
        
        Case 0
        
            Dim ps1 As Integer

            If Not EnableTransparanty(frmPlayer.hWnd, 0) = 0 Then
                'MsgBox "Error using Transparety:" & Chr(10) & "Kein Windows 2000 erkannt!" & Chr(10) & Chr(10) & "Showing form without an effect!", vbCritical
                'frmEinstellungen.Show
                'frmPlayer.Show
            Else
                frmPlayer.Enabled = False
                'frmPlayer.Show
                DoEvents
    
            For ps1 = 200 To 255 Step 3
                DoEvents
                Call EnableTransparanty(frmPlayer.hWnd, ps1)
                DoEvents
            Next
                frmPlayer.Enabled = True
        
            End If
            
        End Select
        
End Sub

Private Sub chkWin2000Effects_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    lblToolTip.Caption = "Aktiviert oder deaktiviert die Windows 2000 Effekte. Einige Fenster werden über den Alphakanal von Windows 2000 eingeblendet."
    
End Sub

Private Sub cmdHideSystray_Click()
    
    SysIcon.HideIcon
    
End Sub

Private Sub cmdOK_Click()
        
    If frmEinstellungen.chkWin2000Effects.Value = 1 Then
    
    'Windows 2000 Effects
    
    Dim ps As Integer

    If Not EnableTransparanty(frmPlayer.hWnd, 0) = 0 Then
        'MsgBox "Error using Transparety:" & Chr(10) & "Kein Windows 2000 erkannt!" & Chr(10) & Chr(10) & "Showing form without an effect!", vbCritical
        'frmEinstellungen.Show
        'frmPlayer.Show
    Else
        frmPlayer.Enabled = False
        'frmPlayer.Show
        DoEvents
    
    For ps = 200 To 255 Step 3
        DoEvents
        Call EnableTransparanty(frmPlayer.hWnd, ps)
        DoEvents
    Next
        frmPlayer.Enabled = True
        
    End If
    End If
    
    frmPlayer.lblEinstellungen.ForeColor = &H808080
    frmPlayer.timerHide.Enabled = True
    frmEinstellungen.Hide
    
End Sub


Private Sub cmdShowSystray_Click()
    
    ' Systray
    
    Set SysIcon = New SYSTRAY
    picSysIcon.picture = ImageListSysIcon.ListImages(SliderSysIcon.Value).picture
    SysIcon.Initialize hWnd, picSysIcon.picture, "sonic"
    SysIcon.ShowIcon
    
End Sub


Private Sub Form_Load()
    
    FramePlayer.Top = FrameStartUp.Top
    FrameInfo.Top = FrameStartUp.Top
    
    frmEinstellungen.Width = 6585
    frmEinstellungen.Height = 6045
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
  Dim msgCallBackMessage As Long
    
  msgCallBackMessage = X / Screen.TwipsPerPixelX
   
  Select Case msgCallBackMessage
    Case WM_MOUSEMOVE
      '"Mouse is moving"
    Case WM_LBUTTONDOWN
      '"Left button went down"
    Case WM_LBUTTONUP
      '"Left button came up"
    Case WM_LBUTTONDBLCLK
        
        frmEinstellungen.txtSystrayIcon.Text = "Hide"
        
        If frmEinstellungen.chkWin2000Effects.Value = 1 Then
        
        'Windows 2000 Effects
    
        Dim ps As Integer

        If Not EnableTransparanty(Me.hWnd, 0) = 0 Then
            'MsgBox "Error using Transparety:" & Chr(10) & "Kein Windows 2000 erkannt!" & Chr(10) & Chr(10) & "Showing form without an effect!", vbCritical
            Me.Show
        Else
            Me.Enabled = False
            DoEvents
    
        For ps = 255 To 0 Step -13
            DoEvents
            Call EnableTransparanty(Me.hWnd, ps)
            DoEvents
        Next
            Me.Enabled = True
        
        End If
        End If
        
        Unload frmOwnMnu
        Unload frmPlayer
        
    Case WM_RBUTTONDOWN
      '"Right button went down"
    Case WM_RBUTTONUP
      '"Right button came up"
    Case WM_RBUTTONDBLCLK
      '"Double click catched from right button"
    Case WM_MBUTTONDOWN
      '"Middle button went down"
    Case WM_MBUTTONUP
      '"Middle button came up"
    Case WM_MBUTTONDBLCLK
      '"Double click catched from middle button"
  End Select
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
        
    SysIcon.HideIcon
    
End Sub




Private Sub FramePlayer_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    lblToolTip.Caption = "sonic"
    
End Sub

Private Sub FrameStartUp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    lblToolTip.Caption = "sonic"
    
End Sub



Private Sub lblHideTime_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    lblToolTip.Caption = "Nach wieviel Sekunden das Player Fenster wieder am linken Bildschirmrand andocken soll."
    
End Sub

Private Sub lblHttp_Click()
    
    'Shell ("explorer mailto: Defcon2@gmx.de")
    Shell ("explorer http:\\www.defcon2.de")
    
End Sub



Private Sub lblSystrayIcon_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    lblToolTip.Caption = "Ändert das Icon in der Systray."
    
End Sub


Private Sub picSysIcon_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    lblToolTip.Caption = "Ändert das Icon in der Systray."
    
End Sub

Private Sub SliderSysIcon_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    lblToolTip.Caption = "Ändert das Icon in der Systray."
    
End Sub

Private Sub SliderSysIcon_Scroll()
    
    SysIcon.HideIcon
    picSysIcon.picture = ImageListSysIcon.ListImages(SliderSysIcon.Value).picture
    SysIcon.Initialize hWnd, picSysIcon.picture, "sonic"
    SysIcon.ShowIcon
        
End Sub



Private Sub TabStripEinstellungen_Click()
    
    Select Case TabStripEinstellungen.SelectedItem.Tag
        Case "Start"
            FrameStartUp.Visible = True
            FramePlayer.Visible = False
            FrameInfo.Visible = False
                        
        Case "Player"
            FrameStartUp.Visible = False
            FramePlayer.Visible = True
            FrameInfo.Visible = False
            
        Case "Info"
            FrameStartUp.Visible = False
            FramePlayer.Visible = False
            FrameInfo.Visible = True
            
    End Select
    
End Sub

Private Sub txtHideTime_KeyPress(KeyAscii As Integer)
        
    If KeyAscii < 49 Or KeyAscii > 57 Then KeyAscii = 0
        
End Sub



Private Sub txtSystrayIcon_Change()
    
    Select Case txtSystrayIcon.Text
        Case "Show"
            cmdShowSystray_Click
        Case "Hide"
            cmdHideSystray_Click
    End Select
    
End Sub
