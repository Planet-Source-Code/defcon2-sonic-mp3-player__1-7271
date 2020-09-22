VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'Kein
   Caption         =   "Form1"
   ClientHeight    =   2625
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4020
   LinkTopic       =   "Form1"
   ScaleHeight     =   2625
   ScaleWidth      =   4020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Visible         =   0   'False
   Begin VB.PictureBox picBeta 
      Appearance      =   0  '2D
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   2520
      Picture         =   "frmSplash.frx":0000
      ScaleHeight     =   330
      ScaleWidth      =   900
      TabIndex        =   4
      Top             =   360
      Width           =   900
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
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
      Left            =   2640
      TabIndex        =   3
      Top             =   2040
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtDeveloper 
      Appearance      =   0  '2D
      BackColor       =   &H00000000&
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
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "developed by Defcon2"
      Top             =   1080
      Width           =   2055
   End
   Begin VB.TextBox txtVersion 
      Appearance      =   0  '2D
      BackColor       =   &H00000000&
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "Version "
      Top             =   1560
      Width           =   1455
   End
   Begin sonic.CLASS_Spriter CLASS_Spriter 
      Height          =   2625
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4020
      _ExtentX        =   7091
      _ExtentY        =   4630
      MaskColor       =   16777215
      MaskPicture     =   "frmSplash.frx":0FBA
      MousePointer    =   0
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
    
    cmdOK.Visible = False
    frmSplash.Hide
    
End Sub
