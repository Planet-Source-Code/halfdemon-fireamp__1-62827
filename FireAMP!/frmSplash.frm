VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   6675
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7035
   LinkTopic       =   "Form1"
   ScaleHeight     =   6675
   ScaleWidth      =   7035
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   5040
      Top             =   4560
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3000
      Left            =   0
      Picture         =   "frmSplash.frx":0000
      ScaleHeight     =   200
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   250
      TabIndex        =   0
      Top             =   0
      Width           =   3750
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   1320
         TabIndex        =   1
         Top             =   2160
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer, j As Integer
Dim s As String
Private Sub Form_Load()
i = 1
s = Space(3) & "<" & Space(3)
Width = Picture1.Width
Height = Picture1.Height
Dim WindowRegion As Long
    WindowRegion = MakeRegion(Picture1)
    SetWindowRgn Me.hwnd, WindowRegion, True
    frmFirePL.Visible = False
End Sub

Private Sub Timer1_Timer()
Label1.Caption = Mid(s, i, Len(s)) & Mid(s, 1, i)
i = i + 1
If i > Len(s) Then
i = 1
j = j + 1

If j > 4 Then
 Unload Me
 frmFireMain.Show
 End If
End If
End Sub
