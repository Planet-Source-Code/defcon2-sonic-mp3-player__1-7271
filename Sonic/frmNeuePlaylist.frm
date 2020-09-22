VERSION 5.00
Begin VB.Form frmNeuePlaylist 
   BorderStyle     =   3  'Fester Dialog
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4065
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
   ScaleHeight     =   2025
   ScaleWidth      =   4065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.Frame FrameNeuePlaylist 
      Height          =   400
      Left            =   20
      TabIndex        =   4
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
         TabIndex        =   5
         Top             =   145
         Width           =   255
      End
      Begin VB.Label lblNeuePlaylistCaption 
         Caption         =   "Neue Playliste erstellen:"
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
         TabIndex        =   6
         Top             =   135
         Width           =   3615
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Abbrechen"
      Height          =   285
      Left            =   3000
      TabIndex        =   3
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox txtNeuePlaylistName 
      Appearance      =   0  '2D
      BorderStyle     =   0  'Kein
      Height          =   255
      Left            =   840
      TabIndex        =   0
      Top             =   1200
      Width           =   3200
   End
   Begin VB.Label lblNeuePlaylist 
      Caption         =   "Einfach den Name der neuen Playliste eingeben und dann im Player auswählen."
      Height          =   495
      Left            =   840
      TabIndex        =   7
      Top             =   480
      Width           =   3015
   End
   Begin VB.Image imgNeuePlaylist 
      Height          =   480
      Left            =   120
      Picture         =   "frmNeuePlaylist.frx":0000
      Top             =   480
      Width           =   480
   End
   Begin VB.Line Line1 
      X1              =   45
      X2              =   4020
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label lblNeuePlaylistName 
      Caption         =   " Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   615
   End
End
Attribute VB_Name = "frmNeuePlaylist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    
    Unload frmNeuePlaylist
    
End Sub

Private Sub cmdClose_Click()
    
    Unload frmNeuePlaylist
    
End Sub

Private Sub cmdOK_Click()
    
    Dim Filename  As String, KeyName As String
        
    frmPlayer.txtPListCount.Text = frmPlayer.txtPListCount.Text + 1
    
    frmPlayer.imgcmbPlaylist.ComboItems.Add , App.Path & "\Playlist" & frmPlayer.txtPListCount.Text & ".pls", txtNeuePlaylistName.Text, 1, 1
    frmPlayer.lstPlaylist.AddItem App.Path & "\Playlist" & frmPlayer.txtPListCount.Text & ".pls"
    
    CreateMP3Playlist App.Path & "\Playlist" & frmPlayer.txtPListCount.Text & ".pls"
    
    FileCopy App.Path & "\Playlist" & frmPlayer.txtPListCount.Text & ".pls", App.Path & "\Playlist" & frmPlayer.txtPListCount.Text & ".pls" & "t"
    
    Filename = App.Path & "\sonic.ini"
    KeyName = "List" & frmPlayer.txtPListCount.Text
    
    WritePrivateProfileString "Playlist", KeyName, txtNeuePlaylistName.Text, Filename
    
    Unload frmNeuePlaylist
    
End Sub

Sub CreateMP3Playlist(Liste As String)
        
    Open Liste For Output As #1
    Print #1, "MP3 Playlist"
    Close #1
    
End Sub



Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    txtNeuePlaylistName.BorderStyle = 0
    lblNeuePlaylistName.BackColor = &H8000000F
    lblNeuePlaylistName.ForeColor = &H80000012
    
End Sub

Private Sub lblNeuePlaylistCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    ReleaseCapture
    SendMessage frmNeuePlaylist.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
    
End Sub

Private Sub lblNeuePlaylistName_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    txtNeuePlaylistName.SetFocus
    txtNeuePlaylistName.BorderStyle = 1
    lblNeuePlaylistName.BackColor = &H8000000D
    lblNeuePlaylistName.ForeColor = &HFFFFFF
    
End Sub

Private Sub txtNeuePlaylistName_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        
        Dim Filename  As String, KeyName As String
        
        frmPlayer.txtPListCount.Text = frmPlayer.txtPListCount.Text + 1
    
        frmPlayer.imgcmbPlaylist.ComboItems.Add , App.Path & "\Playlist" & frmPlayer.txtPListCount.Text & ".pls", txtNeuePlaylistName.Text, 1, 1
        frmPlayer.lstPlaylist.AddItem App.Path & "\Playlist" & frmPlayer.txtPListCount.Text & ".pls"
    
        CreateMP3Playlist App.Path & "\Playlist" & frmPlayer.txtPListCount.Text & ".pls"
    
        FileCopy App.Path & "\Playlist" & frmPlayer.txtPListCount.Text & ".pls", App.Path & "\Playlist" & frmPlayer.txtPListCount.Text & ".pls" & "t"
    
        Filename = App.Path & "\sonic.ini"
        KeyName = "List" & frmPlayer.txtPListCount.Text
    
        WritePrivateProfileString "Playlist", KeyName, txtNeuePlaylistName.Text, Filename
    
        Unload frmNeuePlaylist
        
    End If
    
End Sub

Private Sub txtNeuePlaylistName_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    
    txtNeuePlaylistName.SetFocus
    txtNeuePlaylistName.BorderStyle = 1
    lblNeuePlaylistName.BackColor = &H8000000D
    lblNeuePlaylistName.ForeColor = &HFFFFFF
    
End Sub
