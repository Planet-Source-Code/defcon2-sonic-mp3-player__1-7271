VERSION 5.00
Begin VB.Form frmMsgbox 
   BorderStyle     =   3  'Fester Dialog
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
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
      Height          =   285
      Left            =   3600
      TabIndex        =   4
      Top             =   1920
      Width           =   975
   End
   Begin VB.Frame FrameMsgBox 
      Height          =   400
      Left            =   20
      TabIndex        =   0
      Top             =   -80
      Width           =   4635
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
         Left            =   4320
         TabIndex        =   1
         Top             =   145
         Width           =   255
      End
      Begin VB.Label lblMsgBoxCaption 
         Caption         =   "MsgBoxCaption:"
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
         TabIndex        =   2
         Top             =   135
         Width           =   3615
      End
   End
   Begin VB.Label lblMsgBoxInfo 
      Caption         =   "SONIC Player"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   4565
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Image imgMsgbox 
      Height          =   480
      Left            =   120
      Picture         =   "frmMsgbox.frx":0000
      Top             =   600
      Width           =   480
   End
   Begin VB.Label lblMsgBoxText 
      Caption         =   "MsgboxText..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   975
      Left            =   840
      TabIndex        =   3
      Top             =   600
      Width           =   3615
   End
End
Attribute VB_Name = "frmMsgbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    
    Unload frmMsgbox
    
End Sub

Private Sub cmdOK_Click()
    
    Unload frmMsgbox
    
End Sub

