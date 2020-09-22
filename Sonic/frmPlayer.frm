VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPlayer 
   BorderStyle     =   0  'Kein
   Caption         =   "sonic"
   ClientHeight    =   7995
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   4020
   LinkTopic       =   "Form1"
   Picture         =   "frmPlayer.frx":0000
   ScaleHeight     =   7995
   ScaleWidth      =   4020
   StartUpPosition =   3  'Windows-Standard
   Visible         =   0   'False
   Begin VB.PictureBox picMain 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   8535
      Left            =   0
      OLEDropMode     =   1  'Manuell
      Picture         =   "frmPlayer.frx":0342
      ScaleHeight     =   8535
      ScaleWidth      =   4335
      TabIndex        =   1
      Top             =   0
      Width           =   4335
      Begin VB.CommandButton Command4 
         Caption         =   "k"
         Height          =   255
         Left            =   2520
         TabIndex        =   44
         Top             =   2880
         Width           =   255
      End
      Begin VB.CommandButton Command1 
         Caption         =   "n"
         Height          =   255
         Left            =   2760
         TabIndex        =   43
         Top             =   2880
         Width           =   255
      End
      Begin VB.ListBox lstPlaylist 
         Appearance      =   0  '2D
         Height          =   3930
         Left            =   0
         TabIndex        =   42
         Top             =   2640
         Width           =   375
      End
      Begin VB.TextBox txtPListCount 
         Appearance      =   0  '2D
         Height          =   285
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   480
         Width           =   375
      End
      Begin VB.ListBox lstFileNames 
         Appearance      =   0  '2D
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3930
         Left            =   1920
         TabIndex        =   5
         Top             =   3240
         Width           =   1455
      End
      Begin VB.FileListBox fileFileNames 
         Appearance      =   0  '2D
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3930
         Left            =   0
         TabIndex        =   6
         Top             =   3240
         Width           =   375
      End
      Begin MSComCtl2.FlatScrollBar FlatScrollBar 
         Height          =   3960
         Left            =   3825
         TabIndex        =   32
         Top             =   3210
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   6985
         _Version        =   393216
         LargeChange     =   10
         Max             =   1
         Orientation     =   8323072
      End
      Begin VB.PictureBox picLine 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   3960
         Left            =   3795
         Picture         =   "frmPlayer.frx":23580
         ScaleHeight     =   3960
         ScaleWidth      =   30
         TabIndex        =   4
         Top             =   3210
         Width           =   30
      End
      Begin VB.PictureBox picAdd 
         Appearance      =   0  '2D
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   840
         ScaleHeight     =   735
         ScaleWidth      =   2295
         TabIndex        =   7
         Top             =   4560
         Visible         =   0   'False
         Width           =   2295
         Begin MSComctlLib.ProgressBar AddBar 
            Height          =   255
            Left            =   0
            TabIndex        =   8
            Top             =   480
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   0
            Min             =   1e-4
            Scrolling       =   1
         End
         Begin VB.Label lblAdd 
            BackStyle       =   0  'Transparent
            Caption         =   " Hinzufügen 0/0"
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
            Left            =   0
            TabIndex        =   10
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label lblFolder 
            BackStyle       =   0  'Transparent
            Caption         =   " Ordner"
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
            Left            =   0
            TabIndex        =   9
            Top             =   0
            Width           =   1575
         End
      End
      Begin VB.CommandButton Command3 
         Caption         =   "b"
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
         Left            =   3240
         TabIndex        =   34
         ToolTipText     =   "browse Folder"
         Top             =   2880
         Width           =   255
      End
      Begin VB.ListBox lstFiles 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   3960
         Left            =   480
         OLEDropMode     =   1  'Manuell
         TabIndex        =   33
         Top             =   3210
         Width           =   3015
      End
      Begin VB.CommandButton Command2 
         Caption         =   "c"
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
         TabIndex        =   28
         ToolTipText     =   "clear list"
         Top             =   2880
         Width           =   255
      End
      Begin VB.PictureBox picPauseON 
         Appearance      =   0  '2D
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   960
         Picture         =   "frmPlayer.frx":2464E
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   27
         ToolTipText     =   "Pause..."
         Top             =   2760
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox picPauseOFF 
         Appearance      =   0  '2D
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   960
         Picture         =   "frmPlayer.frx":24BD0
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   26
         Top             =   2760
         Width           =   315
      End
      Begin VB.PictureBox picNextON 
         Appearance      =   0  '2D
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2040
         Picture         =   "frmPlayer.frx":25152
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   25
         ToolTipText     =   "Next..."
         Top             =   2760
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox picNextOFF 
         Appearance      =   0  '2D
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2040
         Picture         =   "frmPlayer.frx":256D4
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   24
         Top             =   2760
         Width           =   315
      End
      Begin VB.PictureBox picBackON 
         Appearance      =   0  '2D
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1680
         Picture         =   "frmPlayer.frx":25C56
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   23
         ToolTipText     =   "Back..."
         Top             =   2760
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox picBackOFF 
         Appearance      =   0  '2D
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1680
         Picture         =   "frmPlayer.frx":261D8
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   22
         Top             =   2760
         Width           =   315
      End
      Begin VB.PictureBox picStopON 
         Appearance      =   0  '2D
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1320
         Picture         =   "frmPlayer.frx":2675A
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   21
         ToolTipText     =   "Stop..."
         Top             =   2760
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox picStopOFF 
         Appearance      =   0  '2D
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1320
         Picture         =   "frmPlayer.frx":26CDC
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   20
         Top             =   2760
         Width           =   315
      End
      Begin VB.PictureBox picPlayON 
         Appearance      =   0  '2D
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   420
         Picture         =   "frmPlayer.frx":2725E
         ScaleHeight     =   495
         ScaleWidth      =   525
         TabIndex        =   19
         ToolTipText     =   "Play..."
         Top             =   2505
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.PictureBox picPlayOFF 
         Appearance      =   0  '2D
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   420
         Picture         =   "frmPlayer.frx":2808C
         ScaleHeight     =   495
         ScaleWidth      =   525
         TabIndex        =   18
         Top             =   2505
         Width           =   525
      End
      Begin VB.TextBox txtCurrentTitle 
         Appearance      =   0  '2D
         Height          =   285
         Left            =   3600
         TabIndex        =   17
         Top             =   7560
         Width           =   375
      End
      Begin VB.PictureBox picAnzeige 
         Appearance      =   0  '2D
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   480
         OLEDropMode     =   1  'Manuell
         ScaleHeight     =   705
         ScaleWidth      =   2625
         TabIndex        =   11
         Top             =   1440
         Width           =   2655
         Begin VB.Label lblTitle 
            Appearance      =   0  '2D
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   " sonic --- MP3 Player for Win32 ..."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   195
            Left            =   120
            TabIndex        =   35
            Top             =   480
            Width           =   2340
         End
         Begin VB.Label lblTime 
            BackStyle       =   0  'Transparent
            Caption         =   "[00]  00:00"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   375
            Left            =   360
            TabIndex        =   16
            Top             =   45
            Width           =   1575
         End
         Begin VB.Label lblVolume 
            BackStyle       =   0  'Transparent
            Caption         =   "100%"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   255
            Left            =   2235
            TabIndex        =   15
            Top             =   120
            Width           =   375
         End
         Begin VB.Label lblVolumeLogo 
            BackStyle       =   0  'Transparent
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "Webdings"
               Size            =   9.75
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   255
            Left            =   2100
            TabIndex        =   14
            Top             =   60
            Width           =   135
         End
         Begin VB.Label lblStatIcon 
            BackStyle       =   0  'Transparent
            Caption         =   "²"
            BeginProperty Font 
               Name            =   "Webdings"
               Size            =   12
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   255
            Left            =   45
            TabIndex        =   13
            Top             =   45
            Width           =   255
         End
         Begin VB.Label lblTime2 
            BackStyle       =   0  'Transparent
            Caption         =   "[00] -00:00"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   375
            Left            =   360
            OLEDropMode     =   1  'Manuell
            TabIndex        =   12
            Top             =   45
            Visible         =   0   'False
            Width           =   1695
         End
      End
      Begin VB.PictureBox picControls 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   0
         Picture         =   "frmPlayer.frx":28EBA
         ScaleHeight     =   65
         ScaleMode       =   0  'Benutzerdefiniert
         ScaleWidth      =   268
         TabIndex        =   3
         Top             =   0
         Width           =   4020
         Begin VB.Image imgPos 
            Height          =   735
            Left            =   3600
            MouseIcon       =   "frmPlayer.frx":2D708
            MousePointer    =   99  'Benutzerdefiniert
            OLEDropMode     =   1  'Manuell
            ToolTipText     =   "Hide..."
            Top             =   240
            Width           =   330
         End
         Begin VB.Image imgExitON 
            Height          =   360
            Left            =   3120
            Picture         =   "frmPlayer.frx":2DA12
            ToolTipText     =   "Beenden..."
            Top             =   480
            Visible         =   0   'False
            Width           =   360
         End
         Begin VB.Image imgExitOFF 
            Height          =   360
            Left            =   3120
            OLEDropMode     =   1  'Manuell
            Picture         =   "frmPlayer.frx":2E114
            Top             =   480
            Width           =   360
         End
         Begin VB.Image imgInfoON 
            Height          =   360
            Left            =   2880
            OLEDropMode     =   1  'Manuell
            Picture         =   "frmPlayer.frx":2E816
            ToolTipText     =   "Info..."
            Top             =   480
            Visible         =   0   'False
            Width           =   360
         End
         Begin VB.Image imgInfoOFF 
            Height          =   360
            Left            =   2880
            OLEDropMode     =   1  'Manuell
            Picture         =   "frmPlayer.frx":2EF18
            Top             =   480
            Width           =   360
         End
      End
      Begin VB.TextBox txtHide 
         Appearance      =   0  '2D
         Height          =   285
         Left            =   3720
         TabIndex        =   2
         Text            =   "0"
         Top             =   7200
         Width           =   255
      End
      Begin VB.Timer timerHide 
         Interval        =   1000
         Left            =   3600
         Top             =   2760
      End
      Begin MSComctlLib.ImageList ImageList 
         Left            =   0
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPlayer.frx":2F61A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPlayer.frx":2FBB6
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageCombo imgcmbPlaylist 
         Height          =   330
         Left            =   480
         TabIndex        =   29
         Top             =   7240
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   14737632
         BackColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImageList       =   "ImageList"
      End
      Begin MSComctlLib.Slider Slider 
         Height          =   255
         Left            =   1080
         TabIndex        =   30
         Top             =   2295
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   450
         _Version        =   393216
         Enabled         =   0   'False
         LargeChange     =   0
         Max             =   1
         TickStyle       =   3
         TickFrequency   =   10
      End
      Begin MSComctlLib.Slider SliderVolume 
         Height          =   615
         Left            =   3120
         TabIndex        =   31
         Top             =   1560
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   1085
         _Version        =   393216
         Orientation     =   1
         Max             =   100
         TickStyle       =   3
      End
      Begin VB.Label lblShuffle 
         BackStyle       =   0  'Transparent
         Caption         =   "Shuffle"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   165
         Left            =   1800
         TabIndex        =   40
         ToolTipText     =   "Shuffle..."
         Top             =   1215
         Width           =   375
      End
      Begin VB.Label lblRepeat 
         BackStyle       =   0  'Transparent
         Caption         =   "Loop"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   165
         Left            =   2280
         TabIndex        =   39
         ToolTipText     =   "Loop..."
         Top             =   1215
         Width           =   375
      End
      Begin VB.Label lblVideo 
         BackStyle       =   0  'Transparent
         Caption         =   "V"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   160
         Left            =   1440
         TabIndex        =   38
         ToolTipText     =   "Video Fenster..."
         Top             =   1220
         Width           =   135
      End
      Begin VB.Label lblEinstellungen 
         BackStyle       =   0  'Transparent
         Caption         =   "Einstellungen"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   165
         Left            =   2640
         TabIndex        =   37
         ToolTipText     =   "Einstellungen..."
         Top             =   1220
         Width           =   735
      End
      Begin VB.Label lblOnTop 
         BackStyle       =   0  'Transparent
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   160
         Left            =   1320
         TabIndex        =   36
         ToolTipText     =   "Im Vordergrund..."
         Top             =   1220
         Width           =   135
      End
      Begin VB.Image imgNoDrag 
         Height          =   6975
         Left            =   3550
         Top             =   960
         Width           =   495
      End
   End
   Begin sonic.CLASS_Spriter CLASS_Spriter 
      Height          =   7995
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4020
      _ExtentX        =   7091
      _ExtentY        =   14102
      MaskPicture     =   "frmPlayer.frx":30152
   End
End
Attribute VB_Name = "frmPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Filename As String
Dim MPlaylist As String
Dim VPlaylist As String
Dim PlayList As String
Dim Pos  As Boolean


Private Sub CLASS_Spriter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        
    Unload frmEditor
        
    ReleaseCapture
    SendMessage frmPlayer.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
    
End Sub



Private Sub CLASS_Spriter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        
    ImgOFF
    txtHide.Text = "0"
            
    frmEinstellungen.txtPlayerPosX.Text = frmPlayer.Top
    frmEinstellungen.txtPlayerPosY.Text = frmPlayer.Left
    
End Sub






Private Sub Command1_Click()
    
    frmNeuePlaylist.Show
    AlwaysOnTop frmNeuePlaylist, 1
    
End Sub

Private Sub Command2_Click()
        
    lstFileNames.Clear
    lstFiles.Clear
    
End Sub

Private Sub Command3_Click()
    
    frmPlayer.timerHide.Enabled = False
    
    Dim tmpPath As String
    tmpPath = frmEinstellungen.txtLastBrowseFolder.Text
    
    tmpPath = BrowseForFolder(tmpPath)
    If tmpPath = "" Then
        
    Else
        
        fileFileNames.Path = tmpPath
        frmEinstellungen.txtLastBrowseFolder.Text = tmpPath
        
            Dim i2 As Integer
            For i2 = 0 To fileFileNames.ListCount - 1
                DoEvents
                picAdd.Visible = True
                lblFolder.Caption = " Ordner " & i & "/" & numFiles
                lblAdd.Caption = " Hinzufügen " & i2 & "/" & fileFileNames.ListCount
                AddBar.Min = 0
                AddBar.Max = fileFileNames.ListCount
                AddBar.Value = i2
                
            Ext = GetExtension(fileFileNames.List(i2))
            
            Select Case Ext
                Case "mp3", "MP3", "Mp3", "mpg", "MPG", "Mpg", "avi", "AVI", "Avi", "mpeg", "MPEG", "Mpeg", "asf", "ASF", "Asf"
                Call ReadMP3(fileFileNames.List(i2), True, True)
                Artist = GetMP3Info.Artist
                Songname = GetMP3Info.Songname
                'Artist = "Artist"
                'Songname = "Songname"
                MP3TAG = lstFiles.ListCount + 1 & ". " & Artist & " - " & Songname
                If Artist = "" Then
                MP3TAG = lstFiles.ListCount + 1 & ". " & fileFileNames.List(i2)
                End If
                lstFiles.AddItem MP3TAG
                lstFileNames.AddItem tmpPath & "\" & fileFileNames.List(i2)
                FlatScrollBar.Max = lstFiles.ListCount - 1
            Case Else
            End Select
            Next i2
        
    End If
        
    frmPlayer.timerHide.Enabled = True
    frmPlayer.SliderVolume.Value = frmEinstellungen.txtVol.Text
    picAdd.Visible = False
    
End Sub

Private Sub Command4_Click()
            
    If txtPListCount.Text = "1" Then
    ShowMsgBox "Hinweis:", "Sie können nicht die letzte Playliste löschen."
    Exit Sub
    End If
    
    On Error GoTo ErrorHandle
    imgcmbPlaylist.ComboItems.Remove imgcmbPlaylist.SelectedItem.index
    lstPlaylist.RemoveItem lstPlaylist.ListIndex
    
    txtPListCount.Text = txtPListCount.Text - 1
    imgcmbPlaylist.ComboItems(1).Selected = True
    
ErrorHandle:
    Select Case Err.Number
        Case 91
            ShowMsgBox "Fehler:", "Bitte wählen Sie erst eine Playliste aus."
            imgcmbPlaylist.ComboItems(1).Selected = True
    End Select
End Sub

Private Sub FlatScrollBar_Scroll()
    
    txtHide.Text = "0"
    lstFiles.Selected(FlatScrollBar.Value) = True
    
End Sub

Private Sub Form_Load()
        
    On Error GoTo ErrorHandle
       

       
    ' ## Variablen
        
    Filename = App.Path & "\sonic.ini"
        
    ' ## Überprüfe die INI Datei
        
    If FileExists(Filename) = 0 Then
    CreateINIFile
    End If
        
    ' ## Load Settings
    
    frmEinstellungen.chkSplash.Value = GetFromInI("Settings", "Splash", Filename)
    
    frmSplash.txtVersion.Text = frmSplash.txtVersion.Text & GetFromInI("Settings", "Version", Filename)
    frmEinstellungen.txtPlayerPosX.Text = GetFromInI("Settings", "PPX", Filename)
    frmEinstellungen.txtPlayerPosY.Text = GetFromInI("Settings", "PPY", Filename)
    frmEinstellungen.chkOnTop.Value = GetFromInI("Settings", "OnTop", Filename)
    frmPlayer.SliderVolume.Value = GetFromInI("Settings", "Vol", Filename)
    frmEinstellungen.txtVol.Text = frmPlayer.SliderVolume.Value
    frmEinstellungen.chkWin2000Effects.Value = GetFromInI("Settings", "Win2000", Filename)
    frmEinstellungen.chkDefault.Value = GetFromInI("Settings", "Default", Filename)
    frmEinstellungen.chkRepeat.Value = GetFromInI("Settings", "Repeat", Filename)
    frmEinstellungen.chkShuffle.Value = GetFromInI("Settings", "Shuffle", Filename)
    frmEinstellungen.SliderSysIcon.Value = GetFromInI("Settings", "Icon", Filename)
    frmEinstellungen.txtLastBrowseFolder.Text = GetFromInI("Settings", "Browse", Filename)
    
    frmEinstellungen.txtVideoPosX.Text = GetFromInI("Video", "PVX", Filename)
    frmEinstellungen.txtVideoPosY.Text = GetFromInI("Video", "PVY", Filename)
    frmEinstellungen.chkVideoOnTop.Value = GetFromInI("Video", "OnTop", Filename)
    frmEinstellungen.chkVideo.Value = GetFromInI("Video", "Show", Filename)
    frmEinstellungen.txtVideoH.Text = GetFromInI("Video", "VideoH", Filename)
    frmEinstellungen.txtVideoW.Text = GetFromInI("Video", "VideoW", Filename)
    
    frmEinstellungen.chkPin.Value = GetFromInI("Settings", "Pin", Filename)
    frmEinstellungen.txtHideTime.Text = GetFromInI("Settings", "HideTime", Filename)
        
    frmEinstellungen.txtPlaylistSelected.Text = GetFromInI("Playlist", "List", Filename)
    frmEinstellungen.chkSaveLastList.Value = GetFromInI("Playlist", "Save", Filename)
    frmEinstellungen.txtListCount.Text = GetFromInI("Playlist", "ListCount", Filename)
    
    ' ## Ende Load Settings
        
        
    ' ## TEST Multiplaylisten
    
    Dim PListPath As String
    
    PListCount = GetFromInI("Playlist", "Count", Filename)
    txtPListCount.Text = PListCount
        
    For l = 1 To PListCount
            
    PListPath = GetFromInI("Playlist", "Path" & l, Filename)
    lstPlaylist.AddItem PListPath
    PListname = GetFromInI("Playlist", "List" & l, Filename)
    imgcmbPlaylist.ComboItems.Add , PListPath, PListname, 1, 1
    PListPathTemp = PListPath & "t"
    
    ' ## Überprüfe die Playlist Dateien
       
    If FileExists(PListPath) = 0 Then
    CreateMP3Playlist (PListPath)
    End If
    
    FileCopy PListPath, PListPathTemp
    Next l
    
        
    
        
    ' ## Farben & Objecte
        
    frmPlayer.CLASS_Spriter.MaskColor = RGB(255, 0, 0)
    frmSplash.CLASS_Spriter.MaskColor = RGB(255, 0, 0)
    
    frmPlayer.lstFiles.ForeColor = RGB(18, 142, 228)
    frmPlayer.imgcmbPlaylist.ForeColor = RGB(18, 142, 228)
    frmPlayer.lblTime.ForeColor = RGB(18, 142, 228)
    frmPlayer.lblTime2.ForeColor = RGB(18, 142, 228)
    frmPlayer.lblVolume.ForeColor = RGB(18, 142, 228)
    
    frmPlayer.fileFileNames.Left = -5000
    frmPlayer.lstFileNames.Left = -5000
    frmPlayer.txtHide.Left = -5000
    frmPlayer.txtCurrentTitle.Left = -5000
    frmPlayer.lstPlaylist.Left = -5000
    frmPlayer.txtPListCount.Left = -5000
    
    ' ## Start
       
    ' ## Playlist
            
    Dim PlsInput As String, PlsInputB As String, PlsInputInA As String, PlsInputInB As String
        
    PlayList = frmEinstellungen.txtPlaylistSelected.Text & "t"
            
    ' ## Auswahl der letzten Playlist
            
    Open PlayList For Input As #1
    
    Input #1, PlsInput
        
    Do While Not EOF(1)
    Input #1, PlsInput, PlsInputB
    
        If PlsInput = "" Then Exit Do
        PlsInputInA = SignToKomma(PlsInput)
        PlsInputInB = SignToKomma(PlsInputB)
        lstFileNames.AddItem PlsInputInA
        lstFiles.AddItem PlsInputInB
        'Call ReadMP3((PlsInput), True, True)
        '    Artist = GetMP3Info.Artist
        '    Songname = GetMP3Info.Songname
        '    MP3Tag = lstFiles.ListCount + 1 & ". " & Artist & " - " & Songname
        '    If Artist = "" Then
        '    MP3Tag = lstFiles.ListCount + 1 & ". " & PlsInput
        '    End If
        '    lstFiles.AddItem MP3Tag
            FlatScrollBar.Max = lstFiles.ListCount - 1
           
    Loop
    
    Close #1
    
    imgcmbPlaylist.ComboItems(Val(frmEinstellungen.txtListCount.Text)).Selected = True
    lstPlaylist.Selected(Val(frmEinstellungen.txtListCount.Text - 1)) = True
    
    ' ## Playlisten kopieren Temp
    
    'FileCopy MPlaylist, App.Path & "\mtemp.mpl"
    'FileCopy VPlaylist, App.Path & "\vtemp.vpl"
    
    ' ## Ende Playlist
    
    frmPlayer.Top = frmEinstellungen.txtPlayerPosX.Text
    frmPlayer.Left = frmEinstellungen.txtPlayerPosY.Text
        
    frmVideo.Top = frmEinstellungen.txtVideoPosX.Text
    frmVideo.Left = frmEinstellungen.txtVideoPosY.Text
    
    frmVideo.Height = frmEinstellungen.txtVideoH.Text
    frmVideo.Width = frmEinstellungen.txtVideoW.Text
    
    Select Case frmEinstellungen.chkSplash.Value ' Abfrage vom Splash Screen
        Case 0
        
        Case 1
            frmSplash.Show
            Pause 1.5
            frmSplash.Hide
    End Select
    
    Select Case frmEinstellungen.chkVideo.Value ' Abfrage vom Video Screen
        Case 0
        
        Case 1
            frmVideo.Show
    End Select
    
    Select Case frmEinstellungen.chkVideoOnTop  ' Abfrage vom Video Screen Immer im Vordergrund
        Case 0
            AlwaysOnTop frmVideo, 0
        Case 1
            AlwaysOnTop frmVideo, 1
    End Select
    
    Select Case frmEinstellungen.chkOnTop.Value ' Abfrage von Immer im Vordergrund
        Case 0
            AlwaysOnTop frmPlayer, 0
            lblOnTop.ForeColor = &H808080
        Case 1
            AlwaysOnTop frmPlayer, 1
            lblOnTop.ForeColor = &HE0E0E0
    End Select
    
    ' ## Systray
    
    frmEinstellungen.txtSystrayIcon.Text = "Show"
        
    ' ## Lautstärke
        
    Dim Vol As Integer
    Vol = (SliderVolume.Value + 1) * 35
    Vol = "-" & Vol
    SliderVolume.Text = 100 - SliderVolume.Value & "%"
    lblVolume.Caption = SliderVolume.Text
    frmVideo.MediaPlayer.Volume = Vol
    
    ' Ende Lautstärke
        
    'Windows 2000 Effects
    
    If frmEinstellungen.chkWin2000Effects.Value = 1 Then
    
    Dim ps As Integer

    If Not EnableTransparanty(Me.hWnd, 0) = 0 Then
        ShowMsgBox "Error using Transparety:", "Kein Windows 2000 erkannt! Bitte überprüfen Sie Ihre Einstellungen. Showing form without an effect!"
        frmEinstellungen.Show
        Me.Show
    Else
        Me.Enabled = False
        Me.Show
        DoEvents
    
    For ps = 0 To 255 Step 15
        DoEvents
        Call EnableTransparanty(Me.hWnd, ps)
        DoEvents
    Next
        Me.Enabled = True
        
    End If
    End If
    
    frmPlayer.Show
    Pos = True
    Exit Sub
    
ErrorHandle:
    ShowMsgBox "Fehler " & Err.Number & ":", Err.Description & "." & (Chr(13)) & "Sollte der Fehler weiter auftreten löschen Sie bitte die sonic.ini im Sonic Verzeichnis."
    frmPlayer.Show
    frmEinstellungen.chkDefault.Value = 1
    Close #1
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    ReleaseCapture
    SendMessage frmPlayer.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    frmEinstellungen.txtPlayerPosX.Text = frmPlayer.Top
    frmEinstellungen.txtPlayerPosY.Text = frmPlayer.Left
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
            
    WritePrivateProfileString "Settings", "Version", App.Major & "." & App.Minor & App.Revision, Filename
    WritePrivateProfileString "Settings", "Splash", CStr(frmEinstellungen.chkSplash.Value), Filename
    
    WritePrivateProfileString "Settings", "PPX", frmEinstellungen.txtPlayerPosX.Text, Filename
    WritePrivateProfileString "Settings", "PPY", frmEinstellungen.txtPlayerPosY.Text, Filename
    WritePrivateProfileString "Settings", "OnTop", CStr(frmEinstellungen.chkOnTop.Value), Filename
    WritePrivateProfileString "Settings", "Vol", CStr(frmPlayer.SliderVolume.Value), Filename
    WritePrivateProfileString "Settings", "Win2000", CStr(frmEinstellungen.chkWin2000Effects.Value), Filename
    WritePrivateProfileString "Settings", "Default", CStr(frmEinstellungen.chkDefault.Value), Filename
    WritePrivateProfileString "Settings", "Repeat", CStr(frmEinstellungen.chkRepeat.Value), Filename
    WritePrivateProfileString "Settings", "Shuffle", CStr(frmEinstellungen.chkShuffle.Value), Filename
    WritePrivateProfileString "Settings", "Icon", CStr(frmEinstellungen.SliderSysIcon.Value), Filename
    WritePrivateProfileString "Settings", "Browse", frmEinstellungen.txtLastBrowseFolder.Text, Filename
    
    WritePrivateProfileString "Video", "PVX", frmEinstellungen.txtVideoPosX.Text, Filename
    WritePrivateProfileString "Video", "PVY", frmEinstellungen.txtVideoPosY.Text, Filename
    WritePrivateProfileString "Video", "OnTop", CStr(frmEinstellungen.chkVideoOnTop.Value), Filename
    WritePrivateProfileString "Video", "Show", CStr(frmEinstellungen.chkVideo.Value), Filename
    WritePrivateProfileString "Video", "VideoH", frmEinstellungen.txtVideoH.Text, Filename
    WritePrivateProfileString "Video", "VideoW", frmEinstellungen.txtVideoW.Text, Filename
    
    WritePrivateProfileString "Settings", "Pin", CStr(frmEinstellungen.chkPin.Value), Filename
    WritePrivateProfileString "Settings", "HideTime", frmEinstellungen.txtHideTime.Text, Filename
    
    WritePrivateProfileString "Playlist", "Save", CStr(frmEinstellungen.chkSaveLastList.Value), Filename
    WritePrivateProfileString "Playlist", "List", frmEinstellungen.txtPlaylistSelected.Text, Filename
    WritePrivateProfileString "Playlist", "ListCount", frmEinstellungen.txtListCount.Text, Filename
    WritePrivateProfileString "Playlist", "Count", frmPlayer.txtPListCount.Text, Filename
    
    ' ## Playlist
    
    If frmEinstellungen.chkSaveLastList.Value = 1 Then
        
    imgcmbPlaylist_Click
    If imgcmbPlaylist.Text = "" Then imgcmbPlaylist.Text = "Playliste"
    
    For pl = 0 To txtPListCount.Text - 1
    
    Dim Pls As String, Pls2 As String
           
    Pls = lstPlaylist.List(pl) & "t"
    Pls2 = lstPlaylist.List(pl)
    
    FileCopy Pls, Pls2
    Kill Pls
        
    Next pl
    
    End If
    
    ' ## Ende Playlist
    
    End
        
End Sub



Private Sub imgcmbPlaylist_Click()
             
    frmEinstellungen.txtPlaylistSelected.Text = imgcmbPlaylist.SelectedItem.Key
    frmEinstellungen.txtListCount.Text = imgcmbPlaylist.SelectedItem.index
        
    ' ## Temp Liste erstellen
    
    Dim PlsTemp As String, PlsTempB As String, PlsTempInA As String, PlsTempInB As String
    
    PlsTemp = lstPlaylist.List(lstPlaylist.ListIndex)
    PlsTemp = PlsTemp & "t"
    
    Open PlsTemp For Output As #1
    Print #1, imgcmbPlaylist.Text
    For p = 0 To lstFileNames.ListCount
    PlsTempInA = KommaToSign(lstFileNames.List(p))
    PlsTempInB = KommaToSign(lstFiles.List(p))
    Print #1, PlsTempInA & "," & PlsTempInB
    Next p
    Close #1
    
    lstPlaylist.Selected(imgcmbPlaylist.SelectedItem.index - 1) = True
    
    ' ## Temp Liste Einlesen
    
    lstFiles.Clear
    lstFileNames.Clear
    
    PlsTemp = lstPlaylist.List(lstPlaylist.ListIndex)
    PlsTemp = PlsTemp & "t"
    
    Open PlsTemp For Input As #1
    
    Input #1, PlsTemp
        
    Do While Not EOF(1)
    Input #1, PlsTemp, PlsTempB
    
        If PlsTemp = "" Then Exit Do
        PlsTempInA = SignToKomma(PlsTemp)
        PlsTempInB = SignToKomma(PlsTempB)
        lstFileNames.AddItem PlsTempInA
        lstFiles.AddItem PlsTempInB
        'Call ReadMP3((PlsTemp), True, True)
        '    Artist = GetMP3Info.Artist
        '    Songname = GetMP3Info.Songname
        '    MP3Tag = lstFiles.ListCount + 1 & ". " & Artist & " - " & Songname
        '    If Artist = "" Then
        '    MP3Tag = lstFiles.ListCount + 1 & ". " & PlsTemp
        '    End If
        '    lstFiles.AddItem MP3Tag
            FlatScrollBar.Max = lstFiles.ListCount - 1
            
    Loop
    
    Close #1
    
    
End Sub



Private Sub imgcmbPlaylist_KeyPress(KeyAscii As Integer)
    
    Dim KeyName As String
    
    Filename = App.Path & "\sonic.ini"
    KeyName = "List" & frmEinstellungen.txtListCount.Text
    
    If KeyAscii = 13 Then
    imgcmbPlaylist.ComboItems(Val(frmEinstellungen.txtListCount.Text)).Text = imgcmbPlaylist.Text
    imgcmbPlaylist.ComboItems(Val(frmEinstellungen.txtListCount.Text)).Selected = True
    WritePrivateProfileString "Playlist", KeyName, imgcmbPlaylist.Text, Filename
    End If
        
End Sub

Private Sub imgExitOFF_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    ImgOFF
    imgExitOFF.Visible = False
    imgExitON.Visible = True
    
End Sub

Private Sub imgExitON_Click()
        
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
        
End Sub



Private Sub imgInfoOFF_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    ImgOFF
    
    imgInfoOFF.Visible = False
    imgInfoON.Visible = True
    
End Sub

Private Sub imgInfoON_Click()
    
    
    frmSplash.Show
    AlwaysOnTop frmSplash, 1
    frmSplash.cmdOK.Visible = True
    'Pause 2
    'frmSplash.Hide
    
End Sub



Private Sub imgPos_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    txtHide.Text = "0"
    
    Select Case Pos
        Case False
            Do
            frmPlayer.Left = frmPlayer.Left + 150
            Loop Until frmPlayer.Left > frmEinstellungen.txtPlayerPosY.Text
            frmPlayer.Left = frmEinstellungen.txtPlayerPosY.Text
            Pos = True
        Case True
            'Do
            'frmPlayer.Left = frmPlayer.Left - 150
            'Loop Until frmPlayer.Left < -3650
            'frmPlayer.Left = -3650
            'Pos = False
        End Select
    
End Sub




Private Sub imgPos_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    
    txtHide.Text = "0"
    
    Select Case Pos
        Case False
            Do
            frmPlayer.Left = frmPlayer.Left + 150
            Loop Until frmPlayer.Left > frmEinstellungen.txtPlayerPosY.Text
            frmPlayer.Left = frmEinstellungen.txtPlayerPosY.Text
            Pos = True
        Case True
            'Do
            'frmPlayer.Left = frmPlayer.Left - 150
            'Loop Until frmPlayer.Left < -3650
            'frmPlayer.Left = -3650
            'Pos = False
        End Select
    
End Sub

Private Sub lblEinstellungen_Click()
        
    Select Case frmEinstellungen.Visible
        Case 0
            
            frmPlayer.lblEinstellungen.ForeColor = &HE0E0E0
            
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
    
            For ps = 255 To 200 Step -5
                DoEvents
                Call EnableTransparanty(frmPlayer.hWnd, ps)
                DoEvents
            Next
                'frmPlayer.Enabled = True
        
            End If
            End If
            
            frmPlayer.timerHide.Enabled = False
            frmEinstellungen.Show
            AlwaysOnTop frmEinstellungen, 1
            
        Case 1
            
            frmPlayer.lblEinstellungen.ForeColor = &H808080
            
            If frmEinstellungen.chkWin2000Effects.Value = 1 Then
    
            'Windows 2000 Effects
    
            Dim ps2 As Integer

            If Not EnableTransparanty(frmPlayer.hWnd, 0) = 0 Then
                'MsgBox "Error using Transparety:" & Chr(10) & "Kein Windows 2000 erkannt!" & Chr(10) & Chr(10) & "Showing form without an effect!", vbCritical
                'frmEinstellungen.Show
                'frmPlayer.Show
            Else
                frmPlayer.Enabled = False
                'frmPlayer.Show
                DoEvents
    
            For ps2 = 200 To 255 Step 3
                DoEvents
                Call EnableTransparanty(frmPlayer.hWnd, ps2)
                DoEvents
            Next
                frmPlayer.Enabled = True
        
            End If
            End If
       
            frmPlayer.timerHide.Enabled = True
            frmEinstellungen.Hide
        
    End Select
    
End Sub

Private Sub lblRepeat_Click()
    
    Select Case frmEinstellungen.chkRepeat.Value
        Case 0
            frmEinstellungen.chkRepeat.Value = 1
            lblRepeat.ForeColor = &HE0E0E0
        Case 1
            frmEinstellungen.chkRepeat.Value = 0
            lblRepeat.ForeColor = &H808080
    End Select
    
End Sub

Private Sub lblOnTop_Click()
    
    Select Case frmEinstellungen.chkOnTop.Value
        Case 0
            frmEinstellungen.chkOnTop.Value = 1
            lblOnTop.ForeColor = &HE0E0E0
        Case 1
            frmEinstellungen.chkOnTop.Value = 0
            lblOnTop.ForeColor = &H808080
    End Select
    
End Sub

Private Sub lblShuffle_Click()
    
    Select Case frmEinstellungen.chkShuffle.Value
        Case 0
            frmEinstellungen.chkShuffle.Value = 1
            lblShuffle.ForeColor = &HE0E0E0
        Case 1
            frmEinstellungen.chkShuffle.Value = 0
            lblShuffle.ForeColor = &H808080
    End Select
    
End Sub

Private Sub lblVideo_Click()
    
    Select Case frmEinstellungen.chkVideo.Value
        Case 0
            frmEinstellungen.chkVideo.Value = 1
            lblVideo.ForeColor = &HE0E0E0
        Case 1
            frmEinstellungen.chkVideo.Value = 0
            lblVideo.ForeColor = &H808080
    End Select
    
End Sub

Private Sub lstFiles_Click()
                
    Dim SEL As Integer
    
    SEL = lstFiles.ListIndex
    lstFileNames.Selected(SEL) = True
    
    FlatScrollBar.Value = lstFiles.ListIndex
    
End Sub

Private Sub imgPos_Click()
    
    txtHide.Text = "0"
    
    Select Case Pos
        Case False
            Do
            frmPlayer.Left = frmPlayer.Left + 250
            Loop Until frmPlayer.Left > frmEinstellungen.txtPlayerPosY.Text
            frmPlayer.Left = frmEinstellungen.txtPlayerPosY.Text
            Pos = True
        Case True
            Do
            frmPlayer.Left = frmPlayer.Left - 250
            Loop Until frmPlayer.Left < -3650
            frmPlayer.Left = -3650
            Pos = False
        End Select
    
    
End Sub


Private Sub lstFiles_DblClick()
        
    Dim File, Tag As String
    File = frmPlayer.lstFileNames.List(frmPlayer.lstFileNames.ListIndex)
        
    frmPlayer.Slider.Enabled = True
    frmPlayer.txtCurrentTitle.Text = frmPlayer.lstFiles.ListIndex + 1
    
    frmVideo.MediaPlayer.Open File
    frmVideo.Caption = File
    frmVideo.timerPlayer.Enabled = True
    frmPlayer.Slider.Min = frmVideo.MediaPlayer.SelectionStart
    
    ' ## Anzeige
    frmPlayer.lblStatIcon.Caption = "4"
    frmPlayer.lblTitle.Caption = " " & frmPlayer.lstFiles.List(frmPlayer.lstFiles.ListIndex)
    
    
    ' ## Extension
    Ext = GetExtension(lstFileNames.List(lstFileNames.ListIndex))
    
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


Private Sub lstFiles_KeyPress(KeyAscii As Integer)
                
    If KeyAscii = 13 Then
    
    Dim File, Tag As String
    File = frmPlayer.lstFileNames.List(frmPlayer.lstFileNames.ListIndex)
        
    frmPlayer.Slider.Enabled = True
    frmPlayer.txtCurrentTitle.Text = frmPlayer.lstFiles.ListIndex + 1
    
    frmVideo.MediaPlayer.Open File
    frmVideo.Caption = File
    frmVideo.timerPlayer.Enabled = True
    frmPlayer.Slider.Min = frmVideo.MediaPlayer.SelectionStart
    
    ' ## Anzeige
    frmPlayer.lblStatIcon.Caption = "4"
    frmPlayer.lblTitle.Caption = " " & frmPlayer.lstFiles.List(frmPlayer.lstFiles.ListIndex)
    
    
    ' ## Extension
    Ext = GetExtension(lstFileNames.List(lstFileNames.ListIndex))
    
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
    End If
    
End Sub

Private Sub lstFiles_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        
    Unload frmEditor
        
    If lstFileNames.List(lstFileNames.ListIndex) = "" Then
    Exit Sub
    End If
        
    If Button = 2 Then
                
        X = GetX * 15.075
        Y = GetY * 15.075
        frmEinstellungen.txtMenuX.Text = X
        frmEinstellungen.txtMenuY.Text = Y
        
        frmPlayer.PopupMenu frmOwnMnu.mnuFile
        
    End If
    
End Sub

Private Sub lstFiles_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    ImgOFF
    
End Sub


Private Sub lstFiles_Scroll()
    
   'FlatScrollBar.Value = lstFiles.ListCount + 1
    
End Sub

Private Sub picAnzeige_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    txtHide.Text = "0"
    ImgOFF
    
End Sub


Private Sub picBackOFF_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    ImgOFF
    picBackOFF.Visible = False
    picBackON.Visible = True
    
End Sub

Private Sub picBackON_Click()
    
    If frmPlayer.lstFiles.Selected(0) = True Then
    frmPlayer.lstFiles.Selected(frmPlayer.lstFiles.ListCount - 1) = True
    picPlayON_Click
    Exit Sub
    End If
    frmPlayer.lstFiles.Selected(frmPlayer.lstFiles.ListIndex - 1) = True
    picPlayON_Click
    
End Sub

Private Sub picControls_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    lstFiles.SetFocus
    ReleaseCapture
    SendMessage frmPlayer.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
    
End Sub

Private Sub picControls_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    ImgOFF
    txtHide.Text = "0"
            
    frmEinstellungen.txtPlayerPosX.Text = frmPlayer.Top
    frmEinstellungen.txtPlayerPosY.Text = frmPlayer.Left
    
End Sub





Private Sub picMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
                            
    lstFiles.SetFocus
    ReleaseCapture
    SendMessage frmPlayer.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
    
    
    
End Sub

Private Sub picMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    ImgOFF
        
End Sub



Private Sub picMain_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
   Ext = GetExtension(Data.Files(1))
            
            Select Case Ext
                Case "mp3", "MP3", "Mp3", "mpg", "MPG", "Mpg", "avi", "AVI", "Avi", "mpeg", "MPEG", "Mpeg", "asf", "ASF", "Asf"
                    Call ReadMP3(Data.Files(1), True, True)
                    Artist = GetMP3Info.Artist
                    Songname = GetMP3Info.Songname
                    MP3TAG = lstFiles.ListCount + 1 & ". " & Artist & " - " & Songname
                    If Artist = "" Then
                    MP3TAG = lstFiles.ListCount + 1 & ". " & Data.Files(1)
                    End If
                    lstFiles.AddItem MP3TAG
                    lstFileNames.AddItem Data.Files(1)
                    FlatScrollBar.Max = lstFiles.ListCount - 1
                Case Else
            End Select
              
        frmPlayer.SliderVolume.Value = frmEinstellungen.txtVol.Text
        picAdd.Visible = False
        
        lstFiles.Selected(lstFiles.ListCount - 1) = True
        frmPlayer.txtCurrentTitle.Text = frmPlayer.lstFiles.ListIndex + 1
        picPlayON_Click
        
End Sub

Private Sub picMain_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    
    If Data.GetFormat(vbCFFiles) Then
        Effect = vbDropEffectCopy
    Else
        Effect = vbDropEffectNone
    End If
    
End Sub

Private Sub picNextOFF_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    ImgOFF
    picNextOFF.Visible = False
    picNextON.Visible = True
    
End Sub



Private Sub picNextON_Click()
                
    If frmPlayer.lstFiles.ListIndex + 1 = frmPlayer.lstFiles.ListCount Then
    frmPlayer.lstFiles.Selected(0) = True
    picPlayON_Click
    Exit Sub
    End If
    
    If frmEinstellungen.chkShuffle.Value = 1 Then
    ShuffleFile = Int(((frmPlayer.lstFiles.ListCount - 1) * Rnd) + 1)
    frmPlayer.lstFiles.Selected(ShuffleFile) = True
    Else
    frmPlayer.lstFiles.Selected(frmPlayer.lstFiles.ListIndex + 1) = True
    End If
        
    picPlayON_Click
    
End Sub

Private Sub picPauseOFF_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    ImgOFF
    picPauseOFF.Visible = False
    picPauseON.Visible = True
    
End Sub

Private Sub picPauseON_Click()
    
    Select Case lblStatIcon.Caption
        Case "4"
            frmVideo.MediaPlayer.Pause
            lblStatIcon.Caption = "1"
        Case Else
            frmVideo.MediaPlayer.Play
            lblStatIcon.Caption = "4"
    End Select
    
End Sub

Private Sub picPlayOFF_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        
    ImgOFF
    picPlayOFF.Visible = False
    picPlayON.Visible = True
    
End Sub



Private Sub picPlayON_Click()
       
    If txtCurrentTitle.Text = "" Then lstFiles.Selected(0) = True
        
    Dim File, Tag As String
    File = frmPlayer.lstFileNames.List(frmPlayer.lstFileNames.ListIndex)
        
    frmPlayer.Slider.Enabled = True
    frmPlayer.txtCurrentTitle.Text = frmPlayer.lstFiles.ListIndex + 1
    
    frmVideo.MediaPlayer.Open File
    frmVideo.Caption = File
    frmVideo.timerPlayer.Enabled = True
    frmPlayer.Slider.Min = frmVideo.MediaPlayer.SelectionStart
    
    ' ## Anzeige
    frmPlayer.lblStatIcon.Caption = "4"
    frmPlayer.lblTitle.Caption = " " & frmPlayer.lstFiles.List(frmPlayer.lstFiles.ListIndex)
    
    
    ' ## Extension
    Ext = GetExtension(lstFileNames.List(lstFileNames.ListIndex))
    
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

Private Sub picStopOFF_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    ImgOFF
    picStopOFF.Visible = False
    picStopON.Visible = True
    
End Sub

Private Sub picStopON_Click()
    
    frmVideo.MediaPlayer.Stop
    frmVideo.Visible = False
    frmVideo.timerPlayer.Enabled = False
    
    ' ## Anzeige
    lblStatIcon.Caption = "²"
    lblTime.Caption = "[00]  00:00"
    lblTime2.Caption = "[00] -00:00"
    lblTitle.Caption = " sonic --- MP3 Player for Win32 ..."
    
End Sub

Private Sub Slider_Click()
    
    frmVideo.MediaPlayer.CurrentPosition = Slider.Value
    
End Sub

Private Sub Slider_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    frmVideo.timerPlayer.Enabled = False
    
End Sub

Private Sub Slider_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    txtHide.Text = "0"
    ImgOFF
    
End Sub

Private Sub Slider_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    frmVideo.timerPlayer.Enabled = True
    
End Sub

Private Sub Slider_Scroll()
        
    'frmVideo.MediaPlayer.CurrentPosition = Slider.Value
        
    sPos = Slider.Value
    Min = Int(sPos / 60): Sec = Int(sPos - Min * 60)
    Slide = Format$(Min, "00") + ":" + Format$(Sec, "00")
    
    Slider.Text = Slide
    
End Sub

Private Sub SliderVolume_Change()
    
    Dim Vol As Integer
        
        
    Vol = (SliderVolume.Value + 1) * 35
    Vol = "-" & Vol
   
        
    SliderVolume.Text = 100 - SliderVolume.Value & "%"
    lblVolume.Caption = SliderVolume.Text
        
    frmVideo.MediaPlayer.Volume = Vol
    frmEinstellungen.txtVol.Text = SliderVolume.Value
    
End Sub

Private Sub SliderVolume_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    txtHide.Text = "0"
    ImgOFF
    
End Sub

Private Sub SliderVolume_Scroll()
    
    Dim Vol As Integer
        
        
    Vol = (SliderVolume.Value + 1) * 35
    Vol = "-" & Vol
   
        
    SliderVolume.Text = 100 - SliderVolume.Value & "%"
    lblVolume.Caption = SliderVolume.Text
        
    frmVideo.MediaPlayer.Volume = Vol
    frmEinstellungen.txtVol.Text = SliderVolume.Value
    
End Sub

Private Sub timerHide_Timer()
    
    
    If frmEinstellungen.chkPin.Value = 1 Then
    txtHide.Text = "0"
    Exit Sub
    Else:
    txtHide.Text = txtHide.Text + 1
        If txtHide.Text = frmEinstellungen.txtHideTime.Text Then
        Do
        frmPlayer.Left = frmPlayer.Left - 250
        Loop Until frmPlayer.Left < -3650
        frmPlayer.Left = -3650
        txtHide.Text = "0"
        Pos = False
        
        End If
        
    End If
    
End Sub

Private Sub lstFiles_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
                    
                        
    'Count number of files
    Dim numFiles As Integer, MP3TAG As String
    numFiles = Data.Files.Count

    'Add all dropped files into the list
    Dim i As Integer
    For i = 1 To numFiles
        DoEvents
        'File or directory?
        If (GetAttr(Data.Files(i)) And vbDirectory) = vbDirectory Then
            
            
            fileFileNames.Path = Data.Files(i)
                        
            Dim i2 As Integer
            For i2 = 0 To fileFileNames.ListCount - 1
                DoEvents
                picAdd.Visible = True
                lblFolder.Caption = " Ordner " & i & "/" & numFiles
                lblAdd.Caption = " Hinzufügen " & i2 & "/" & fileFileNames.ListCount
                AddBar.Min = 0
                AddBar.Max = fileFileNames.ListCount
                AddBar.Value = i2
                
                Call ReadMP3(fileFileNames.List(i2), True, True)
                Artist = GetMP3Info.Artist
                Songname = GetMP3Info.Songname
                MP3TAG = lstFiles.ListCount + 1 & ". " & Artist & " - " & Songname
                If Artist = "" Then
                MP3TAG = lstFiles.ListCount + 1 & ". " & fileFileNames.List(i2)
                End If
                lstFiles.AddItem MP3TAG
                lstFileNames.AddItem Data.Files(i) & "\" & fileFileNames.List(i2)
                FlatScrollBar.Max = lstFiles.ListCount - 1
                
            Next i2
            
        Else
                            
            picAdd.Visible = True
            lblFolder.Caption = " Dateien " & numFiles
            lblAdd.Caption = " Hinzufügen " & i & "/" & numFiles
            AddBar.Min = 0
            AddBar.Max = numFiles
            AddBar.Value = i
            
            Ext = GetExtension(Data.Files(i))
            
            Select Case Ext
                Case "mp3", "MP3", "Mp3", "mpg", "MPG", "Mpg", "avi", "AVI", "Avi", "mpeg", "MPEG", "Mpeg", "asf", "ASF", "Asf"
                    Call ReadMP3(Data.Files(i), True, True)
                    Artist = GetMP3Info.Artist
                    Songname = GetMP3Info.Songname
                    MP3TAG = lstFiles.ListCount + 1 & ". " & Artist & " - " & Songname
                    If Artist = "" Then
                    MP3TAG = lstFiles.ListCount + 1 & ". " & Data.Files(i)
                    End If
                    lstFiles.AddItem MP3TAG
                    lstFileNames.AddItem Data.Files(i)
                    FlatScrollBar.Max = lstFiles.ListCount - 1
                Case Else
            End Select
            
        End If
    Next i
        
        frmPlayer.SliderVolume.Value = frmEinstellungen.txtVol.Text
        picAdd.Visible = False
    
End Sub

Private Sub lstFiles_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
        
    If Data.GetFormat(vbCFFiles) Then
        Effect = vbDropEffectCopy
    Else
        Effect = vbDropEffectNone
    End If
    
End Sub

Private Sub lblTime2_Click()
    
    lblTime2.Visible = False
    lblTime.Visible = True
    
End Sub

Private Sub lblTime_Click()
    
    lblTime2.Visible = True
    lblTime.Visible = False
    
End Sub

Private Sub picAnzeige_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
   Ext = GetExtension(Data.Files(1))
            
            Select Case Ext
                Case "mp3", "MP3", "Mp3", "mpg", "MPG", "Mpg", "avi", "AVI", "Avi", "mpeg", "MPEG", "Mpeg", "asf", "ASF", "Asf"
                    Call ReadMP3(Data.Files(1), True, True)
                    Artist = GetMP3Info.Artist
                    Songname = GetMP3Info.Songname
                    MP3TAG = lstFiles.ListCount + 1 & ". " & Artist & " - " & Songname
                    If Artist = "" Then
                    MP3TAG = lstFiles.ListCount + 1 & ". " & Data.Files(1)
                    End If
                    lstFiles.AddItem MP3TAG
                    lstFileNames.AddItem Data.Files(1)
                    FlatScrollBar.Max = lstFiles.ListCount - 1
                Case Else
            End Select
              
        frmPlayer.SliderVolume.Value = frmEinstellungen.txtVol.Text
        picAdd.Visible = False
        
        lstFiles.Selected(lstFiles.ListCount - 1) = True
        frmPlayer.txtCurrentTitle.Text = frmPlayer.lstFiles.ListIndex + 1
        picPlayON_Click
        
End Sub

Private Sub picAnzeige_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    
    If Data.GetFormat(vbCFFiles) Then
        Effect = vbDropEffectCopy
    Else
        Effect = vbDropEffectNone
    End If
    
End Sub

Sub CreateMP3Playlist(Liste As String)
    
    ShowMsgBox "Fehler:", "Keine MP3Playliste gefunden! Erstelle neue Datei."
    
    Open Liste For Output As #1
    Print #1, "MP3 Playlist"
    Close #1
    
End Sub


Sub CreateINIFile()
    
    ShowMsgBox "Fehler:", "Keine INI gefunden! Erstelle neue Datei."
    
    Open App.Path & "\sonic.ini" For Output As #1
    Print #1, "[Settings]"
    Print #1, "Version=" & App.Major & "." & App.Minor & App.Revision
    Print #1, "Splash=1"
    Print #1, "PPX=1000"
    Print #1, "PPY=1000"
    Print #1, "OnTop=0"
    Print #1, "Vol=15"
    Print #1, "Pin=0"
    Print #1, "HideTime=8"
    Print #1, "Win2000=0"
    Print #1, "Default=1"
    Print #1, "Repeat=0"
    Print #1, "Shuffle=0"
    Print #1, "Icon=1"
    Print #1, "Browse=c:\"
    Print #1, "[Video]"
    Print #1, "PVX=2325"
    Print #1, "PVY=4845"
    Print #1, "OnTop=0"
    Print #1, "Show=1"
    Print #1, "VideoH=6000"
    Print #1, "VideoW=8000"
    Print #1, "[Playlist]"
    Print #1, "List=" & App.Path & "\Playlist1.pls"
    Print #1, "ListCount=1"
    Print #1, "Save=0"
    Print #1, "Count=1"
    Print #1, "List1=Playliste 1"
    Print #1, "List2=Playliste 2"
    Print #1, "List3=Playliste 3"
    Print #1, "List4=Playliste 4"
    Print #1, "List5=Playliste 5"
    Print #1, "List6=Playliste 6"
    Print #1, "List7=Playliste 7"
    Print #1, "List8=Playliste 8"
    Print #1, "List9=Playliste 9"
    Print #1, "List10=Playliste 10"
    Print #1, "Path1=" & App.Path & "\Playlist1.pls"
    Print #1, "Path2=" & App.Path & "\Playlist2.pls"
    Print #1, "Path3=" & App.Path & "\Playlist3.pls"
    Print #1, "Path4=" & App.Path & "\Playlist4.pls"
    Print #1, "Path5=" & App.Path & "\Playlist5.pls"
    Print #1, "Path6=" & App.Path & "\Playlist6.pls"
    Print #1, "Path7=" & App.Path & "\Playlist7.pls"
    Print #1, "Path8=" & App.Path & "\Playlist8.pls"
    Print #1, "Path9=" & App.Path & "\Playlist9.pls"
    Print #1, "Path10=" & App.Path & "\Playlist10.pls"
    Close #1
    
    
End Sub

Sub ImgOFF()
    
    txtHide.Text = "0"
        
    imgExitOFF.Visible = True
    imgExitON.Visible = False
    picPlayOFF.Visible = True
    picPlayON.Visible = False
    picStopOFF.Visible = True
    picStopON.Visible = False
    picBackOFF.Visible = True
    picBackON.Visible = False
    picNextOFF.Visible = True
    picNextON.Visible = False
    picPauseOFF.Visible = True
    picPauseON.Visible = False
    imgInfoOFF.Visible = True
    imgInfoON.Visible = False
    
End Sub

Public Function BrowseForFolder(selectedPath As String) As String
Dim Browse_for_folder As BROWSEINFOTYPE
Dim itemID As Long
Dim selectedPathPointer As Long
Dim tmpPath As String * 256
With Browse_for_folder
    .hOwner = Me.hWnd ' Window Handle
    .lpszTitle = "Durchsuchen... " ' Dialog Title
    .lpfn = FunctionPointer(AddressOf BrowseCallbackProcStr) ' Dialog callback function that preselectes the folder specified
    selectedPathPointer = LocalAlloc(LPTR, Len(selectedPath) + 1) ' Allocate a string
    CopyMemory ByVal selectedPathPointer, ByVal selectedPath, Len(selectedPath) + 1 ' Copy the path to the string
    .lParam = selectedPathPointer ' The folder to preselect
End With
itemID = SHBrowseForFolder(Browse_for_folder) ' Execute the BrowseForFolder API
If itemID Then
    If SHGetPathFromIDList(itemID, tmpPath) Then ' Get the path for the selected folder in the dialog
        BrowseForFolder = Left$(tmpPath, InStr(tmpPath, vbNullChar) - 1) ' Take only the path without the nulls
    End If
    Call CoTaskMemFree(itemID) ' Free the itemID
End If
Call LocalFree(selectedPathPointer) ' Free the string from the memory
End Function



