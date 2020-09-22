VERSION 5.00
Begin VB.Form frmFireAuxPL 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form4"
   ClientHeight    =   4920
   ClientLeft      =   1050
   ClientTop       =   825
   ClientWidth     =   4680
   LinkTopic       =   "Form4"
   ScaleHeight     =   4920
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3585
      Left            =   0
      ScaleHeight     =   239
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   187
      TabIndex        =   0
      Top             =   0
      Width           =   2805
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005C5CCD&
         Height          =   2760
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   2295
      End
   End
End
Attribute VB_Name = "frmFireAuxPL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Width = Picture1.Width
Height = Picture1.Height
Dim WindowRegion As Long
    WindowRegion = MakeRegion(Picture1)
    SetWindowRgn Me.hwnd, WindowRegion, True

End Sub


Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
      ReleaseCapture
      SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If
End Sub



