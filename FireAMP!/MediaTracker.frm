VERSION 5.00
Begin VB.Form frmMediaTracker 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "FireAMP-MediaTracker"
   ClientHeight    =   2625
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   2625
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picSkin 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1500
      Left            =   0
      Picture         =   "MediaTracker.frx":0000
      ScaleHeight     =   100
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   200
      TabIndex        =   0
      Top             =   0
      Width           =   3000
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "lblTitle"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0FF&
         Height          =   285
         Left            =   180
         TabIndex        =   3
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label lblTime 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "00:00 [00:00]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0E0FF&
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   960
         Width           =   2415
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FireAMP!         [Media Tracker]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   360
         TabIndex        =   1
         Top             =   180
         Width           =   2445
      End
   End
End
Attribute VB_Name = "frmMediaTracker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

Dim windowRegion As Long
Height = picSkin.Height
Width = picSkin.Width
    windowRegion = MakeRegion(picSkin)
    SetWindowRgn Me.hwnd, windowRegion, True
    
SetWindowPos hwnd, HWND_TOPMOST, _
   0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
   
Top = Screen.Height - Height - 300
Left = Screen.Width - Width
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
SetWindowPos hwnd, HWND_TOPMOST, _
   0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE

End Sub

Private Sub Form_Unload(Cancel As Integer)
SetWindowPos hwnd, HWND_NOTOPMOST, _
   0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE

End Sub

Private Sub picSkin_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 2 Then
Hide

frmFireMain.Show
frmFireMain.WindowState = 0
End If

End Sub
