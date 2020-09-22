VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmFirePL 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   6345
   ClientLeft      =   2070
   ClientTop       =   2055
   ClientWidth     =   10545
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   10545
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3120
      Left            =   240
      Picture         =   "Playlist.frx":0000
      ScaleHeight     =   3120
      ScaleWidth      =   150
      TabIndex        =   4
      Top             =   840
      Width           =   150
      Begin VB.PictureBox picBar 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   10
         Picture         =   "Playlist.frx":1A42
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   8
         TabIndex        =   5
         Top             =   120
         Width           =   120
      End
   End
   Begin VB.Timer tmrMainPosition 
      Interval        =   10
      Left            =   6960
      Top             =   3840
   End
   Begin VB.ListBox lstPaths 
      Height          =   2370
      Left            =   6840
      TabIndex        =   3
      Top             =   1080
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.PictureBox picSkin 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4500
      Left            =   0
      Picture         =   "Playlist.frx":1C04
      ScaleHeight     =   300
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   300
      TabIndex        =   0
      Top             =   0
      Width           =   4500
      Begin MSComctlLib.ListView lstPl 
         Height          =   2835
         Left            =   600
         TabIndex        =   1
         ToolTipText     =   "The Playlist"
         Top             =   960
         Width           =   3300
         _ExtentX        =   5821
         _ExtentY        =   5001
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FlatScrollBar   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   12632319
         BackColor       =   0
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Title"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Length"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lblCaption 
         BackStyle       =   0  'Transparent
         Caption         =   "PlayList"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   600
         TabIndex        =   2
         Top             =   480
         Width           =   2355
      End
   End
End
Attribute VB_Name = "frmFirePL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private yy As Long, Temp As Long
Private Sub Form_Load()
Width = picSkin.Width
Height = picSkin.Height
Dim windowRegion As Long
    windowRegion = MakeRegion(picSkin)
    SetWindowRgn Me.hwnd, windowRegion, True

lstPL.ColumnHeaders.Item(1).Width = CInt(lstPL.Width / 1.3)
lstPL.ColumnHeaders.Item(2).Width = lstPL.Width - lstPL.ColumnHeaders.Item(1).Width


Left = frmFireMain.Left + frmFireMain.Width



End Sub

Private Sub lstPL_Click()
On Error GoTo e:
    picBar.Top = (picBack.ScaleHeight - picBar.Height) * (lstPL.SelectedItem.Index / lstPL.ListItems.Count)
Exit Sub
e:
End Sub

Private Sub lstPl_DblClick()
On Error GoTo e
    frmFireMain.picBtn_Click 1
    curFile = lstPaths.List(lstPL.SelectedItem.Index - 1)
    frmFirePL.lstPaths.ListIndex = lstPL.SelectedItem.Index - 1
    frmFireMain.picBtn_Click 0
    
        
Exit Sub
e:
End Sub

Private Sub lstPl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
On Error Resume Next
If getExtension(lstPL.SelectedItem.Text) <> "mp3" Then
frmDummy.mnuTagEdit.Enabled = False
Else
frmDummy.mnuTagEdit.Enabled = True
End If
        Me.PopupMenu frmDummy.mnuPlaylist
End If
End Sub

Private Sub picBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 1 Then
If picBar.Top + picBar.Height > picBack.ScaleHeight And yy < Y Then Exit Sub
If picBar.Top < 10 And yy > Y Then Exit Sub

If Temp < lstPL.ListItems.Count / 2 Then
Temp = ((picBar.Top) * lstPL.ListItems.Count) / picBack.ScaleHeight

Else
Temp = ((picBar.Top + picBar.Height) * lstPL.ListItems.Count) / picBack.ScaleHeight

End If

lstPL.ListItems(Temp + 1).EnsureVisible


If Abs(yy - Y) > 50 Then Exit Sub
       picBar.Top = picBar.Top + Y
       yy = Y
       
       
End If

End Sub


Private Sub picBar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If picBar.Top + picBar.Height > picBack.ScaleHeight And yy < Y Then picBar.Top = picBack.ScaleHeight - picBar.Height
If picBar.Top < 10 And yy > Y Then picBar.Top = 0

End Sub

Private Sub picSkin_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
      ReleaseCapture
      SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If
End Sub

Private Sub tmrMainPosition_Timer()
Left = (frmFireMain.Left - frmFireMain.Width + 1700)
Top = frmFireMain.Top + 500
End Sub

