VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmFireMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "FireAMP!"
   ClientHeight    =   9465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10725
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00C00000&
   Icon            =   "FireAMP_Main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9465
   ScaleWidth      =   10725
   StartUpPosition =   2  'CenterScreen
   Begin MSScriptControlCtl.ScriptControl sc1 
      Left            =   6240
      Top             =   4440
      _ExtentX        =   1005
      _ExtentY        =   1005
      AllowUI         =   -1  'True
   End
   Begin VB.Timer tmrVis 
      Interval        =   200
      Left            =   240
      Top             =   7200
   End
   Begin VB.PictureBox picCtrlSrc 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   510
      Left            =   6480
      Picture         =   "FireAMP_Main.frx":11C2
      ScaleHeight     =   34
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   34
      TabIndex        =   12
      Top             =   6360
      Visible         =   0   'False
      Width           =   510
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   7680
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Open Media"
      Filter          =   "MIDI Files|*.mid;*.rmi|WAVE files|*.wav|MPEG Audio|*.mp1;*.mp2;*.mp3|MPEG Video|*.mpg;*.mpeg;*.mpe|Any File|*.*"
   End
   Begin VB.PictureBox picBtnSrc 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1920
      Left            =   7560
      Picture         =   "FireAMP_Main.frx":1FD4
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   8
      Top             =   5640
      Width           =   960
   End
   Begin VB.PictureBox ScopeBuff 
      AutoRedraw      =   -1  'True
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C0E0FF&
      Height          =   2010
      Left            =   6360
      ScaleHeight     =   134
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   262
      TabIndex        =   9
      Top             =   360
      Visible         =   0   'False
      Width           =   3930
   End
   Begin VB.Timer Timer_Greq 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1200
      Top             =   7200
   End
   Begin VB.Timer tmrTime 
      Interval        =   100
      Left            =   720
      Top             =   7200
   End
   Begin VB.Timer tmrPbr 
      Interval        =   10
      Left            =   2160
      Top             =   7200
   End
   Begin VB.Timer tmrInfo 
      Interval        =   200
      Left            =   1680
      Top             =   7200
   End
   Begin VB.PictureBox picSkin 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      FillColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5520
      Left            =   0
      Picture         =   "FireAMP_Main.frx":8016
      ScaleHeight     =   368
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   369
      TabIndex        =   0
      Top             =   0
      Width           =   5535
      Begin VB.PictureBox Frame1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1560
         ScaleHeight     =   255
         ScaleWidth      =   3495
         TabIndex        =   17
         Top             =   3960
         Width           =   3495
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "lblInfo"
            ForeColor       =   &H00C0E0FF&
            Height          =   225
            Left            =   0
            TabIndex        =   18
            Top             =   0
            Width           =   495
         End
      End
      Begin VB.PictureBox picBtn 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   465
         Index           =   2
         Left            =   1680
         ScaleHeight     =   465
         ScaleWidth      =   495
         TabIndex        =   11
         Top             =   4500
         Width           =   495
      End
      Begin VB.PictureBox picBtn 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   465
         Index           =   1
         Left            =   1080
         ScaleHeight     =   465
         ScaleWidth      =   495
         TabIndex        =   10
         Top             =   4500
         Width           =   495
      End
      Begin VB.PictureBox picBtn 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   465
         Index           =   0
         Left            =   480
         ScaleHeight     =   465
         ScaleWidth      =   495
         TabIndex        =   7
         Top             =   4500
         Width           =   495
      End
      Begin VB.PictureBox picBarBack 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   150
         Left            =   480
         Picture         =   "FireAMP_Main.frx":6B918
         ScaleHeight     =   10
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   300
         TabIndex        =   5
         Top             =   4200
         Width           =   4500
         Begin VB.PictureBox picBarFront 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   150
            Left            =   0
            Picture         =   "FireAMP_Main.frx":6DC82
            ScaleHeight     =   150
            ScaleWidth      =   150
            TabIndex        =   6
            Top             =   0
            Width           =   150
         End
      End
      Begin VB.PictureBox picCtrl 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   1
         Left            =   4680
         ScaleHeight     =   12
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   12
         TabIndex        =   4
         Top             =   360
         Width           =   180
      End
      Begin VB.PictureBox picCtrl 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   0
         Left            =   4410
         ScaleHeight     =   12
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   12
         TabIndex        =   3
         Top             =   360
         Width           =   180
      End
      Begin VB.PictureBox fraDisplay 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   2895
         Left            =   480
         ScaleHeight     =   2895
         ScaleWidth      =   4575
         TabIndex        =   14
         Top             =   960
         Width           =   4575
         Begin VB.PictureBox Scope 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            FontTransparent =   0   'False
            Height          =   2175
            Left            =   240
            ScaleHeight     =   145
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   273
            TabIndex        =   15
            Top             =   600
            Width           =   4095
            Begin VB.Label lblVis 
               Alignment       =   2  'Center
               BackColor       =   &H005C5CCD&
               BackStyle       =   0  'Transparent
               ForeColor       =   &H00FFFFFF&
               Height          =   345
               Left            =   0
               TabIndex        =   16
               Top             =   0
               Visible         =   0   'False
               Width           =   4035
            End
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "FireAMP!"
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
            Left            =   120
            TabIndex        =   20
            Top             =   0
            Width           =   1020
         End
         Begin VB.Label lblAlbum 
            BackStyle       =   0  'Transparent
            Caption         =   "Burning Media Inc,"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0E0FF&
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   360
            Width           =   4095
         End
      End
      Begin VB.Frame fraVideo 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   3015
         Left            =   480
         TabIndex        =   13
         Top             =   840
         Visible         =   0   'False
         Width           =   4575
      End
      Begin VB.Label lblStatus 
         BackStyle       =   0  'Transparent
         Caption         =   "00:00 [00:00]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0FF&
         Height          =   255
         Left            =   480
         TabIndex        =   2
         Top             =   3960
         Width           =   975
      End
      Begin VB.Label lblCaption 
         BackStyle       =   0  'Transparent
         Caption         =   "FireAMP !"
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
         Height          =   255
         Left            =   600
         TabIndex        =   1
         Top             =   300
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmFireMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
'
'================================================================================
' FireAMP! -- main
'================================================================================
'
'
Private infoStr As String, Pos As Integer, infoStr1 As String
Private Temp As Integer
Private Temp_Str As String
Private isMute As Boolean
Private Const dimColor = &HC0C0C0
Private cur As Integer
Dim t As Long ' debugging
Private visIndex As Integer, St As Integer
Private step As Integer, s1 As Integer, s2 As Integer, s3 As Integer

Private Sub Form_GotFocus()
lblCaption.ForeColor = vbWhite
Unload frmMediaTracker
End Sub

' use true for debug mode: no playlist, fast startup, default file play
' great for testing vis.
#Const DBUG = False
Private Sub Form_Load()


t = Timer
plColor = &HC0C0FF
Dim windowRegion As Long
Divisor = 10
St = 1
infoStr = "   Greetings, User!   *  Welcome to FireAMP!   "
 
Pos = 1


Width = picSkin.Width
Height = picSkin.Height

    windowRegion = MakeRegion(picSkin)
    SetWindowRgn Me.hwnd, windowRegion, True




 

Abt = False
isMute = False

Wx = picBtnSrc.ScaleWidth / 2
Wy = picBtnSrc.ScaleHeight / 4

cWx = picCtrlSrc.ScaleWidth / 2
cWy = picCtrlSrc.ScaleHeight / 2

Dim i

For i = 0 To 2
picBtn(i).Width = Wx
picBtn(i).Height = Wy
Next

For i = 0 To 1
picCtrl(i).Width = cWx
picCtrl(i).Height = cWy
Next

BitBlt picBtn(0).hdc, 0, 0, Wx, Wy, picBtnSrc.hdc, 0, 0, vbSrcCopy 'play
BitBlt picBtn(1).hdc, 0, 0, Wx, Wy, picBtnSrc.hdc, 0, Wx, vbSrcCopy 'stop
BitBlt picBtn(2).hdc, 0, 0, Wx, Wy, picBtnSrc.hdc, 0, Wx * 3, vbSrcCopy 'open

BitBlt picCtrl(0).hdc, 0, 0, cWx, cWy, picCtrlSrc.hdc, 0, 0, vbSrcCopy
BitBlt picCtrl(1).hdc, 0, 0, cWx, cWy, picCtrlSrc.hdc, 0, cWy, vbSrcCopy

' debug mode
#If DBUG Then
' fill up the path of the clip below
curFile = "D:\ksk new\misc\anime\mp3\Midori no Hibi - Sentimental.mp3"
picBtn_Click 0
Unload frmFirePL
#End If



#If Not DBUG Then
frmFirePL.Show
#End If


Debug.Print String(30, "-")
Debug.Print "FireAMP! - Starting UP!" & vbCrLf
Debug.Print Abs(Timer - t) & " Seconds for start up, yeah!"

visIndex = GetSetting(App.EXEName, "Settings", "Visualization", 0)

dimColor1 = 100
dimColor2 = 100
dimColor3 = 100

stepColor1 = 150
stepColor2 = 100
stepColor3 = 50

End Sub

Private Sub Form_LostFocus()
lblCaption.ForeColor = dimColor
End Sub


Private Sub lblCaption_Click()
PopupMenu frmDummy.mnuFireAMP
End Sub

'seek bar
Private Sub picBarFront_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 And Not (FireAMP_Pos Is Nothing) Then
tmrPbr.Enabled = False
picBarFront.Left = picBarFront.Left + (X / Screen.TwipsPerPixelX)
lblStatus.Caption = "Seeking... " & convertToStdTime(getBarPosition(picBarFront, picBarBack, FireAMP_Pos.Duration))
End If
End Sub

Private Sub picBarFront_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
FireAMP_Pos.CurrentPosition = getBarPosition(picBarFront, picBarBack, FireAMP_Pos.Duration)
tmrPbr.Enabled = True
End Sub

' play, stop and open
Public Sub picBtn_Click(Index As Integer)
On Error Resume Next
Static isPaused As Boolean
Dim LST As ListItem
Dim Count As Integer
Select Case Index
Case 0:
If curFile = "" Then Exit Sub
Dim ext As String
ext = getExtension(curFile)
 If Not (ext = "mpg" Or ext = "mpeg" Or ext = "dat" Or ext = "wmv") Then
      INITIALIZE_GREQ
End If
picBtn_MouseUp 0, 1, 0, 0, 0
infoStr = "   *   FireAMP!   [Playing] " & getFileName(curFile) & "      "

PlayClip curFile

lblStatus.Visible = True
picBarBack.Visible = True

If getExtension(curFile) = "mp3" Then
Dim tempTag As tagMP3ID3V1
tempTag = readMP3Tag(curFile)
lblTitle.Caption = Trim(tempTag.Artist) & " - " & Trim(tempTag.Title)
lblAlbum.Caption = tempTag.Album
infoStr1 = Space(4) & " * " & Space(4) & lblTitle.Caption
infoStr = infoStr & "Clip: " & tempTag.Title & " * Artist: " & tempTag.Artist
infoStr = infoStr & " * Album: " & tempTag.Album & " * Genre: " & getGenre(tempTag.Genre)
Else
Dim parts() As String
parts = Split(getFileName(curFile), ".")
lblTitle.Caption = parts(0)
lblAlbum.Caption = parts(1) & " file"
End If
frmMediaTracker.lblTitle.Caption = lblTitle.Caption

For Count = 1 To frmFirePL.lstPL.ListItems.Count
If frmFirePL.lstPL.ListItems.Item(Count).Bold Then frmFirePL.lstPL.ListItems.Item(Count).Bold = False
frmFirePL.lstPL.ListItems.Item(Count).ForeColor = plColor
 
Next Count
Set LST = frmFirePL.lstPL.ListItems.Item(frmFirePL.lstPL.SelectedItem.Index)
LST.SubItems(1) = convertToStdTime(FireAMP_Pos.Duration)
frmFirePL.lstPL.SelectedItem.Bold = True
LST.EnsureVisible
BitBlt Frame1.hdc, 0, 0, Frame1.ScaleWidth, Frame1.ScaleWidth, picSkin.hdc, Frame1.Left, Frame1.Top, vbSrcCopy

Case 1:
StopClip
DoStop
lblStatus.Caption = "00:00 [00:00]"
lblStatus.Visible = False
picBarBack.Visible = False
fraVideo.Visible = False
fraDisplay.Visible = False
infoStr = "   Greetings, User!   *  Welcome to FireAMP!   "

Case 2:
cd1.FILENAME = ""
cd1.ShowOpen
curFile = cd1.FILENAME
If curFile <> "" Then
StopClip
frmFirePL.lstPaths.AddItem curFile

Set LST = frmFirePL.lstPL.ListItems.Add(, , getFileName(curFile))
Dim Temp_Player As New FilgraphManager
Dim Temp_Pos As IMediaPosition
On Error GoTo e
Temp_Player.RenderFile curFile
Set Temp_Pos = Temp_Player
LST.SubItems(1) = convertToStdTime(Temp_Pos.Duration)

Set Temp_Pos = Nothing
Set Temp_Player = Nothing
 picBtn_Click 0
End If


End Select
picSkin.SetFocus
Exit Sub
e:
LST.SubItems(1) = "??:??"
LST.ForeColor = vbRed
End Sub

Private Sub picBtn_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim offSetY As Integer
Select Case Index
Case 0: 'play
If Not isPlaying Then
offSetY = 0
Else
offSetY = Wx * 2
End If
Case 1: 'stop
offSetY = Wx
Case 2: 'open
offSetY = Wx * 3
End Select

BitBlt picBtn(Index).hdc, 0, 0, Wx, Wy, picBtnSrc.hdc, Wx, offSetY, vbSrcCopy
picBtn(Index).Refresh
End Sub

Private Sub picBtn_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim offSetY As Integer
Select Case Index
Case 0: 'play
If Not isPlaying Then
offSetY = 0
Else
offSetY = Wx * 2
End If
Case 1: 'stop
offSetY = Wx
Case 2: 'open
offSetY = Wx * 3
End Select

BitBlt picBtn(Index).hdc, 0, 0, Wx, Wy, picBtnSrc.hdc, 0, offSetY, vbSrcCopy
picBtn(Index).Refresh

End Sub

Private Sub picCtrl_Click(Index As Integer)
Select Case Index
Case 0:
 Me.WindowState = vbMinimized
frmMediaTracker.Show
Me.Hide
Case 1:
Form_Unload 0
End Select
'picSkin.SetFocus
End Sub

Private Sub picCtrl_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Index
Case 0
BitBlt picCtrl(0).hdc, 0, 0, cWx, cWy, picCtrlSrc.hdc, cWx, 0, vbSrcCopy
Case 1
BitBlt picCtrl(1).hdc, 0, 0, cWx, cWy, picCtrlSrc.hdc, cWx, cWy, vbSrcCopy
End Select
picCtrl(Index).Refresh
End Sub

Private Sub picCtrl_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Index
Case 0
BitBlt picCtrl(0).hdc, 0, 0, cWx, cWy, picCtrlSrc.hdc, 0, 0, vbSrcCopy
Case 1
BitBlt picCtrl(1).hdc, 0, 0, cWx, cWy, picCtrlSrc.hdc, 0, cWy, vbSrcCopy
End Select
picCtrl(Index).Refresh
End Sub

Private Sub picSkin_GotFocus()
lblCaption.ForeColor = vbWhite
End Sub

' keyboard short cuts
Private Sub picSkin_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case Chr(KeyCode)
 Case "X", "x": ' goodbye
Form_Unload 0
 
 Case "F", "f": ' full screen video
 Dim ext As String
 ext = getExtension(curFile)
 If ext = "mpg" Or ext = "mpeg" Or ext = "dat" Or ext = "wmv" Then
 FireAMP_VideoWin.HideCursor True
  refreshVideo frmFullScreen.Frame1
 frmFullScreen.Show
 Me.Hide
 End If
Case " ": ' play
picBtn_Click 0
Case "S", "s": ' stop
picBtn_Click 1
Case "O", "o": ' open
picBtn_Click 2
Case "P", "p": ' show playlist

Case "G", "g": ' volume up
changeVolume True

Case "H", "h": ' volume down
 changeVolume False

Case "M", "m": 'mute
 If isMute Then
  FireAMP_Vol.Volume = 0
lblCaption.Caption = "FireAMP !"
Else
 FireAMP_Vol.Volume = -10000
lblCaption.Caption = "FireAMP ! [Mute]"
End If
isMute = Not isMute

Case "V", "v"
 visIndex = visIndex - 1
 If visIndex < 0 Then visIndex = 16
 dispVis visIndex
Case "B", "b"
 visIndex = visIndex + 1
 If visIndex > 16 Then visIndex = 0
 dispVis visIndex

Case "N", "n":
 picCtrl_Click (0)
Case "C", "c"
 PlayVideoCD
 
End Select


End Sub

Private Sub picSkin_LostFocus()
lblCaption.ForeColor = dimColor
End Sub

Private Sub picSkin_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
      ReleaseCapture
      SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If
frmFirePL.Show
SetFocus
End Sub


Private Sub changeVolume(UPorDOWN As Boolean)
On Error Resume Next

If UPorDOWN Then  ' vol. up
currentVolume = currentVolume - 500
If currentVolume < -5000 Then currentVolume = -5000
FireAMP_Vol.Volume = currentVolume
Else ' vol. down
currentVolume = currentVolume + 500
If currentVolume > 0 Then currentVolume = 0
FireAMP_Vol.Volume = currentVolume
End If
lblStatus.Caption = "Vol: " & (5000 + currentVolume) / 100 * 2 & "%"


End Sub

Private Sub tmrInfo_Timer()

If Len(lblTitle.Caption) > 40 Then
lblTitle.Caption = Mid(infoStr1, Pos, Len(infoStr1)) & Mid(infoStr1, 1, Pos)
End If
lblInfo.Caption = Mid(infoStr, Pos, Len(infoStr)) & Mid(infoStr, 1, Pos)
Pos = Pos + 1
If Pos > Len(infoStr) Then Pos = 1
End Sub

Private Sub tmrPbr_Timer()

If (Not (FireAMP_Pos Is Nothing) And isPlaying) Then
If FireAMP_Pos.Duration = FireAMP_Pos.CurrentPosition And frmFirePL.lstPaths.ListIndex < frmFirePL.lstPaths.ListCount - 1 Then

lblTitle.Caption = "Changing Track..."
lblAlbum.Caption = ""

picBtn_Click 1
curFile = frmFirePL.lstPaths.List(frmFirePL.lstPaths.ListIndex + 1)
frmFirePL.lstPaths.ListIndex = frmFirePL.lstPaths.ListIndex + 1
On Error Resume Next
frmFirePL.lstPL.ListItems(frmFirePL.lstPaths.ListIndex + 1).Selected = True
picBtn_Click 0

End If
On Error Resume Next
picBarFront.Left = (picBarBack.ScaleWidth - picBarFront.Width) * (FireAMP_Pos.CurrentPosition / FireAMP_Pos.Duration)

Else
fraVideo.Visible = False
fraDisplay.Visible = False
DoStop
End If
End Sub

Private Sub tmrTime_Timer()
If Not (FireAMP_Pos Is Nothing) Then
lblStatus.Caption = convertToStdTime(FireAMP_Pos.CurrentPosition) & " [" & convertToStdTime(FireAMP_Pos.Duration) & "]"
frmMediaTracker.lblTime.Caption = lblStatus.Caption
frmMediaTracker.lblTitle.Caption = lblTitle.Caption

Else
fraVideo.Visible = False
fraDisplay.Visible = False
DoStop

End If
End Sub
'===================================================================================

Private Sub Form_Unload(Cancel As Integer)
 Set FSys = Nothing ' release FSys
    If DevHandle <> 0 Then
        Call DoStop
End If
SaveSetting App.EXEName, "Settings", "Visualization", visIndex
End
End Sub

Private Sub Timer_Greq_Timer()
' Timer for greq update
Static buff As String * 255
buff = Space(255)
Static WAVEFORMAT As WaveFormatEx
    With WAVEFORMAT
        .FormatTag = WAVE_FORMAT_PCM
        .Channels = 1
        .SamplesPerSec = 11025 '11khz
        .BitsPerSample = 16
        .BlockAlign = (.Channels * .BitsPerSample) \ 8
        .AvgBytesPerSec = .BlockAlign * .SamplesPerSec
        .ExtraDataSize = 0
    End With
    Debug.Print "waveInOpen:"; mciGetErrorString(waveInOpen(DevHandle, -1, VarPtr(WAVEFORMAT), 0, 0, 0), buff, 255)
    
    Debug.Print vbCrLf & buff
    If DevHandle = 0 Then
        Dim e As ErrStruct
        e.errNum = 2
        e.errShortDesc = "Could not open WaveIn device!"
        e.errLongDesc = "FireAMP! could not open the WaveIn device. This can happen when FireAMP! was not shut down properly or the device is in use by another application. Restart FireAMP! to fix this problem or close the other application"
        logError e
        Timer_Greq.Enabled = False
        Exit Sub
    End If
    Debug.Print " "; DevHandle
    Call waveInStart(DevHandle)
    
    
    Timer_Greq.Enabled = False
    Call Visualize

End Sub

' Initialize graphic equalizer
Public Sub INITIALIZE_GREQ()
    Call DoReverse
    
    ScopeBuff.Width = Scope.Width
    ScopeBuff.Height = Scope.Height
    
    ScopeBuff.BackColor = Scope.BackColor
    ScopeHeight = ScopeBuff.ScaleHeight
    Timer_Greq.Enabled = True
                        Randomize Timer
                         s1 = Rnd * 10 + 1
                         s2 = Rnd * 10 + 1
                         s3 = Rnd * 10 + 1
                         step = Rnd * 20 + 1

'avs style???
' i'am workin' on it
'Dim a As New FireVisualization
'Set a.thePictureBox = ScopeBuff
'sc1.AddObject "scope", a
'sc1.AddCode "const sh =" & ScopeHeight
        
End Sub

Public Sub Visualize()
    'These are all static as they get allocated somewhere and persist for a long time
    'takes the pressure off the heap
    
    Static X As Long ' the main position var., also the array index
    Static Wave As WAVEHDR
    
    ' data to draw the vis. with
    Static InData(0 To NUMSAMPLES - 1) As Integer       ' wave-in data
    Static OutData(0 To NUMSAMPLES - 1) As Single       ' fft out data
    Static PeakData(0 To NUMSAMPLES - 1) As Single      ' peak of a specific frequency
    Static ExData(0 To NUMSAMPLES - 1) As Single        ' peak of peak of a specific frequency
    Static WxData(0 To NUMSAMPLES - 1) As Single        ' peak of wave in-data
    Static WxDataEx(0 To NUMSAMPLES - 1) As Single      ' peak of peak of wave-in data
    Static PeakFallData(0 To NUMSAMPLES - 1) As Integer ' time to fall yet?
    
    
    'some useful vars...
    Static yy As Integer, Col As Long, CX As Integer, CY As Integer
    Static newx As Single, newy As Single
    Static angle As Integer
    Static c As Integer
    Static Temp As Long
    Static dblDegrees As Double, dblRadians As Double
    
    With ScopeBuff 'Save some time referencing it...
    
        Do
        'lpdata requires the address of an array to fill up data with
            Wave.lpData = VarPtr(InData(0))
        'the buffer length
            Wave.dwBufferLength = NUMSAMPLES
        ' ???
            Wave.dwFlags = 0
         'prepare device for input
            Call waveInPrepareHeader(DevHandle, VarPtr(Wave), Len(Wave))
            Call waveInAddBuffer(DevHandle, VarPtr(Wave), Len(Wave))
            
            ' if the following statement is removed, the vis. will be a lot faster (avs style)
            ' but uses up 100% of cpu!
            ' this is why i hate avs
            Sleep 50 ' give device a breather
            
            ' the following loop is quite useless, but anyway...
            Do
                'Just wait for the blocks to be done or the device to close
            Loop Until ((Wave.dwFlags And WHDR_DONE) = WHDR_DONE) Or DevHandle = 0
            
            
            If DevHandle = 0 Then Exit Do 'Cut out if the device is closed
            
            Call waveInUnprepareHeader(DevHandle, VarPtr(Wave), Len(Wave))
            Call FFTAudio(InData, OutData) ' fft the in-data
           
            .Cls
                            
            ScopeBuff.CurrentX = -1
            ScopeBuff.CurrentY = (ScopeHeight / 2)
        
            
            For X = 10 To 255 Step St
             
                
                
                On Error Resume Next
                     If c > 2000 Then
                     Randomize Timer
                        CX = Rnd * 4 + 1
                        CY = Rnd * 4 + 1
                         c = 0
                         step = Rnd * 20 + 1
                         s1 = Rnd * 10 + 1
                         s2 = Rnd * 10 + 1
                         s3 = Rnd * 10 + 1
                     End If
                     
                        c = c + 1
If (PeakData(X * 2) < Abs(OutData(X * 2))) Then PeakData(X * 2) = OutData(X * 2) 'peak of outdata
If (ExData(X * 2) < Abs(PeakData(X * 2))) Then
  ExData(X * 2) = PeakData(X * 2) 'peak of peak data
  PeakFallData(X * 2) = 0 ' reset time to fall
End If

' same as above
If (WxData(X * 2) < Abs(InData(X * 2))) Then WxData(X * 2) = InData(X * 2)
If (WxDataEx(X * 2) < Abs(WxData(X * 2))) Then WxDataEx(X * 2) = WxData(X * 2)

Select Case visIndex

Case 0 ' Fire, revisited

yy = ScopeHeight - Sqr(Abs(PeakData(X * 2)) / Divisor)
ScopeBuff.Line (X, ScopeHeight)-(X, yy), &H80FF&    ' main bars

yy = ScopeHeight - Sqr(Abs(ExData(X * 2)) / Divisor) * 1.1

'embers on fire

ScopeBuff.PSet (X, yy), vbRed
ScopeBuff.PSet (X, yy - 3), RGB(150, 0, 0)
ScopeBuff.PSet (X, yy - 6), RGB(50, 0, 0)


PeakData(X * 2) = PeakData(X * 2) - 15000
ExData(X * 2) = ExData(X * 2) - 2500
    St = 1
    
'------------------------------------------------------------------------------------
Case 1 'bars, revisited


yy = ScopeHeight - Sqr(Abs(PeakData(X * 2)) / Divisor)
ScopeBuff.Line (X, ScopeHeight)-(X + 4, yy), &HC000&, BF ' main bars

yy = ScopeHeight - Sqr(Abs(ExData(X * 2)) / Divisor) * 1.1

' embers on bars
ScopeBuff.Line (X, yy)-(X + 4, yy + 1), &HC00000, BF


PeakData(X * 2) = PeakData(X * 2) - 15000
If PeakFallData(X * 2) >= 5 Then ' is it time fall yet?
ExData(X * 2) = ExData(X * 2) - 2500 'yes? then start falling
End If

PeakFallData(X * 2) = PeakFallData(X * 2) + 1
St = 6
'------------------------------------------------------------------------------------
Case 2 'audio analyser, revisited

Divisor = Divisor * 0.3
yy = Sqr(Abs(PeakData(X * 2))) \ Divisor

For Temp = 0 To yy Step 2
Col = RGB((Temp) Mod dimColor1 + stepColor1, (Temp) Mod dimColor2 + stepColor2, (Temp) Mod dimColor3 + stepColor3)
ScopeBuff.Line (X, ScopeHeight - Temp)-(X + 8, ScopeHeight - Temp), Col

Next Temp
yy = ScopeHeight - Sqr(Abs(ExData(X * 2))) \ Divisor * 0.7

ScopeBuff.Line (X, yy)-(X + 8, yy), RGB(200, 200, 200)

If PeakFallData(X * 2) >= 5 Then
ExData(X * 2) = ExData(X * 2) - 2500
End If

PeakData(X * 2) = PeakData(X * 2) - 5000

PeakFallData(X * 2) = PeakFallData(X * 2) + 1

St = 9
Divisor = Divisor / 0.3

'------------------------------------------------------------------------------------
Case 3 'embers, also revisited

yy = ScopeHeight - Sqr(Abs(PeakData(X * 2)) / Divisor)
ScopeBuff.Line (X, yy)-(X + 4, yy - 1), RGB(150, 0, 0), BF '150,0,0
ScopeBuff.Line (X, yy - 2)-(X + 4, yy - 3), RGB(100, 0, 0), BF '100,0,0
ScopeBuff.Line (X, yy - 4)-(X + 4, yy - 5), RGB(50, 0, 0), BF '50,0,0

PeakData(X * 2) = PeakData(X * 2) - 2500
St = 6
'------------------------------------------------------------------------------------
Case 4 'fire rain, revisited


yy = ScopeHeight - Sqr(Abs(PeakData(X * 2)) / Divisor)
ScopeBuff.Line (X, yy)-(X, yy - 1), RGB(150, 0, 0), BF '150,0,0
ScopeBuff.Line (X, yy - 2)-(X, yy - 3), RGB(100, 0, 0), BF '100,0,0
ScopeBuff.Line (X, yy - 4)-(X, yy - 5), RGB(50, 0, 0), BF '50,0,0

PeakData(X * 2) = PeakData(X * 2) - 2500
St = 1

'------------------------------------------------------------------------------------
Case 5 'pond ripple, revisited

yy = ScopeHeight - Sqr(Abs(PeakData(X * 2)) / Divisor)
ScopeBuff.Circle (ScopeBuff.ScaleWidth / 2, ScopeBuff.ScaleHeight / 2), yy Mod 100, RGB(yy Mod 100, X Mod 100, 100)
PeakData(X * 2) = PeakData(X * 2) - 2500
St = 1

'------------------------------------------------------------------------------------
Case 6 'rain

yy = ScopeHeight - Sqr(Abs(PeakData(X * 2)) / Divisor)
ScopeBuff.Circle (ScopeBuff.ScaleWidth / CX, ScopeBuff.ScaleHeight / CY), yy Mod 100, RGB(yy Mod 100, X Mod 100, 100)
         
ScopeBuff.Circle (ScopeBuff.ScaleWidth / CY, ScopeBuff.ScaleHeight / CX), yy Mod 100, RGB(yy Mod 100, X Mod 100, 100)
          

PeakData(X * 2) = PeakData(X * 2) - 2500
St = 1
'------------------------------------------------------------------------------------
Case 7 ' spike ball

dblDegrees = CDbl((360 / 255) * (X * 2))

' Convert Degrees to Radians
dblRadians = dblDegrees * (3.14159265 / 180)

yy = Sqr(Abs(WxData(X * 2))) \ 3
Col = RGB(yy * 5, yy * 0, yy * 0) 'y*5,0,0
ScopeBuff.Line (ScopeBuff.ScaleWidth / 2, ScopeBuff.ScaleHeight / 2)-((ScopeBuff.ScaleWidth / 2) + (yy) * Cos(dblRadians), (ScopeBuff.ScaleHeight / 2) + (yy) * Sin(dblRadians)), Col
WxData(X * 2) = WxData(X * 2) - 700
St = 1
'------------------------------------------------------------------------------------
Case 8 ' xplosion

dblDegrees = CDbl((360 / 255) * (X * 2))

' Convert Degrees to Radians
dblRadians = dblDegrees * (3.14159265 / 180)

yy = Sqr(Abs(WxData(X * 2))) \ 3
Col = RGB(yy * 10, yy * 5, yy * 2.5)

ScopeBuff.PSet ((ScopeBuff.ScaleWidth / 2) + (yy) * Cos(dblRadians), (ScopeBuff.ScaleHeight / 2) + (yy) * Sin(dblRadians)), Col
WxData(X * 2) = WxData(X * 2) - 1000

'------------------------------------------------------------------------------------
Case 9 'wave train, revisited


yy = ScopeHeight - Sqr(Abs(PeakData(X * 2)) / Divisor)

Col = RGB((ScopeHeight - yy) * 5, 0, 0)
ScopeBuff.Line (X, yy)-(X, ScopeHeight - yy + 1), Col

PeakData(X * 2) = PeakData(X * 2) - 2500

St = 1

'------------------------------------------------------------------------------------
Case 10 'spectroscope, revisited: thanks to punisher

yy = ScopeHeight - Sqr(Abs(PeakData(X * 2))) / 2
yy = Math.Log(X) * Sqr(yy)
Col = RGB(X Mod 100 * 3.141592654, (X * yy) Mod 50 + 100, (yy * X) Mod 3.14152654)

ScopeBuff.Line (X, yy)-(X, ScopeHeight - yy + 1), Col, BF

PeakData(X * 2) = PeakData(X * 2) - 5000
St = 1

'------------------------------------------------------------------------------------
Case 11 'wave - wave

Static Wx As Long, Wy As Long
yy = (ScopeHeight / 2) - Sqr(Abs(WxData(X * 2) / 10))
Col = RGB((ScopeHeight - yy) * 1.7, 0, 0)
ScopeBuff.Line -(X - 10, yy), Col
WxData(X * 2) = WxData(X * 2) - 2500
St = 2

'------------------------------------------------------------------------------------
Case 12 ' abstract - Highway

yy = Sqr(Abs(WxData(X * 2))) / 5
For Temp = 0 To yy Step 2

Col = RGB((yy) * 10, yy * 10, yy * 10) 'yy*10,0,0

ScopeBuff.Line (X, Tan(Temp) * 100)-((X) + 10, Tan(Temp) * 100), Col
ScopeBuff.Line (X, Tan(Temp) * 50)-((X) + 10, Tan(Temp) * 50), Col
  
Next Temp
WxData(X * 2) = WxData(X * 2) - 2500
St = 11
'------------------------------------------------------------------------------------
' the vis below are my original creations
' great for soft music
' these equations have a lot of power
' just mess around with the angle variable
' by default, step is a pseudo-random value
Case 13 ' particle - randomization
ScopeBuff.DrawWidth = 2
yy = Sqr(Abs(WxData(X * 2))) / 3

dblDegrees = CDbl((360 / 255) * (X * 2))
dblRadians = dblDegrees * (3.14159265 / 180)

Col = RGB((yy) * s1, yy * s2, yy * s3) 'yy*10,0,0



newx = Cos(dblDegrees) * angle + Scope.ScaleWidth / 2
newy = Sin(dblDegrees) * yy + ScopeHeight / 2

ScopeBuff.PSet (newx, newy), Col
ScopeBuff.DrawWidth = 1
ScopeBuff.PSet (newx + 3, newy - 3), Col
ScopeBuff.PSet (newx - 3, newy + 3), Col
ScopeBuff.PSet (newx + 3, newy + 3), Col
ScopeBuff.PSet (newx - 3, newy - 3), Col


newx = Cos(dblDegrees) * yy + Scope.ScaleWidth / 2
newy = Sin(dblDegrees) * angle + ScopeHeight / 2

.DrawWidth = 2
ScopeBuff.PSet (newx, newy), Col

ScopeBuff.DrawWidth = 1
ScopeBuff.PSet (newx + 3, newy - 3), Col
ScopeBuff.PSet (newx - 3, newy + 3), Col
ScopeBuff.PSet (newx + 3, newy + 3), Col
ScopeBuff.PSet (newx - 3, newy - 3), Col



WxData(X * 2) = WxData(X * 2) - 2500
St = 11
'comment the following and...
angle = angle + step

'as far i can see..
' angle=angle+1 draws an exploding thingy
' 5 draws a slowly spinning exploding thingy
' 9 draws a slowly spinning imploding thingy
' 21 draws a confusing thingy

'angle = angle + 1 'mess up here
If (angle > 100) Then angle = 0


'----------------------------------------------------------------------------------------
Case 14 'particle - warp
ScopeBuff.DrawWidth = 2
yy = Sqr(Abs(WxData(X * 2))) / 3

dblDegrees = CDbl((360 / 255) * (X * 2))
dblRadians = dblDegrees * (3.14159265 / 180)

Col = RGB((yy) * s1, yy * s2, yy * s3) 'yy*10,0,0



newx = Tan(dblDegrees) * angle + Scope.ScaleWidth / 2
newy = Tan(dblDegrees) * yy + ScopeHeight / 2

ScopeBuff.PSet (newx, newy), Col
ScopeBuff.DrawWidth = 1
ScopeBuff.PSet (newx + 3, newy - 3), Col
ScopeBuff.PSet (newx - 3, newy + 3), Col
ScopeBuff.PSet (newx + 3, newy + 3), Col
ScopeBuff.PSet (newx - 3, newy - 3), Col

newx = Tan(dblDegrees) * yy + Scope.ScaleWidth / 2
newy = Tan(dblDegrees) * angle + ScopeHeight / 2

ScopeBuff.PSet (newx, newy), Col
ScopeBuff.DrawWidth = 1
ScopeBuff.PSet (newx + 3, newy - 3), Col
ScopeBuff.PSet (newx - 3, newy + 3), Col
ScopeBuff.PSet (newx + 3, newy + 3), Col
ScopeBuff.PSet (newx - 3, newy - 3), Col



  
WxData(X * 2) = WxData(X * 2) - 2500
St = 11
'angle = angle + step
angle = angle + 5
If (angle > 100) Then angle = 0

'------------------------------------------------------------------------------------
Case 15 'particle - over head traffic
ScopeBuff.DrawWidth = 2
yy = Sqr(Abs(WxData(X * 2))) / 3

dblDegrees = CDbl((360 / 255) * (X * 2))
dblRadians = dblDegrees * (3.14159265 / 180)

Col = RGB((yy) * s1, yy * s2, yy * s3) 'yy*10,0,0



newx = Atn(dblDegrees) * angle + Scope.ScaleWidth / 2
newy = Tan(dblDegrees) * yy + ScopeHeight / 2

ScopeBuff.PSet (newx, newy), Col
ScopeBuff.DrawWidth = 1
ScopeBuff.PSet (newx + 3, newy - 3), Col
ScopeBuff.PSet (newx - 3, newy + 3), Col
ScopeBuff.PSet (newx + 3, newy + 3), Col
ScopeBuff.PSet (newx - 3, newy - 3), Col

newx = Tan(dblDegrees) * yy + Scope.ScaleWidth / 2
newy = Atn(dblDegrees) * angle + ScopeHeight / 2

ScopeBuff.PSet (newx, newy), Col
ScopeBuff.DrawWidth = 1
ScopeBuff.PSet (newx + 3, newy - 3), Col
ScopeBuff.PSet (newx - 3, newy + 3), Col
ScopeBuff.PSet (newx + 3, newy + 3), Col
ScopeBuff.PSet (newx - 3, newy - 3), Col



  
WxData(X * 2) = WxData(X * 2) - 2500
St = 11
'angle = angle + step
angle = angle + 30
If (angle > 100) Then angle = 0

'------------------------------------------------------------------------------------

Case 16 'particle - spin
ScopeBuff.DrawWidth = 2
yy = Sqr(Abs(WxData(X * 2))) / 3

dblDegrees = CDbl((360 / 255) * (X * 2))
dblRadians = dblDegrees * (3.14159265 / 180)

Col = RGB((yy) * s1, yy * s2, yy * s3) 'yy*10,0,0



newx = Cos(dblDegrees) * angle + Scope.ScaleWidth / 2
newy = Sin(dblDegrees) * yy + ScopeHeight / 2

ScopeBuff.PSet (newx, newy), Col
ScopeBuff.DrawWidth = 1
ScopeBuff.PSet (newx + 3, newy - 3), Col
ScopeBuff.PSet (newx - 3, newy + 3), Col
ScopeBuff.PSet (newx + 3, newy + 3), Col
ScopeBuff.PSet (newx - 3, newy - 3), Col

 
WxData(X * 2) = WxData(X * 2) - 2500
St = 11
'angle = angle + step
angle = angle + 5
If (angle > 100) Then angle = 0
'------------------------------------------------------------------------------------

End Select

            Next X
           ScopeBuff.CurrentY = ScopeBuff.ScaleWidth

            Scope.Picture = .Image 'Display the double-buffer
            
            DoEvents


    
    
           Loop While DevHandle <> 0

    End With

    Visualizing = False
End Sub

Private Sub tmrVis_Timer()
Static i As Integer
i = i + 1
If i > 5 Then
i = 0
tmrVis.Enabled = False
lblVis.Visible = False
End If
End Sub

Sub PlayVideoCD()
curFile = "E:\mpegav\avseq01.dat"
If Not FSys.FileExists(curFile) Then
 Dim e As ErrStruct
 e.errNum = 3
 e.errShortDesc = "Invalid CD in drive"
 e.errLongDesc = "FireAMP! tried to play a video CD but it appears that the CD in the drive is not a videoCD" & vbCrLf & "Replace with a videoCD and try again"
 logError e
 Exit Sub
End If
picBtn_Click 0
End Sub
Sub dispVis(visIndex As Integer)
Select Case visIndex
Case 0
lblVis.Caption = "Primodial - Fire"
Case 1
lblVis.Caption = "Primodial - Bars"
Case 2
lblVis.Caption = "Primodial - Audio Analyser"
Case 3
lblVis.Caption = "Primodial - Embers"
Case 4
lblVis.Caption = "Primodial - Fire Rain"
Case 5
lblVis.Caption = "Wet - Pond Ripple"
Case 6
lblVis.Caption = "Wet - Rain"
Case 7
lblVis.Caption = "Spike - Spike Ball"
Case 8
lblVis.Caption = "Spike - Xplosion"
Case 9
lblVis.Caption = "Waves - Wave Train"
Case 10
lblVis.Caption = "Waves - Spectroscope"
Case 11
lblVis.Caption = "Waves - Wave"
Case 12
lblVis.Caption = "Abstract - Highway"
Case 13
lblVis.Caption = "Particle - Randomization"
Case 14
lblVis.Caption = "Particle - Warp"
Case 15
lblVis.Caption = "Particle - Over Head Traffic"
Case 16
lblVis.Caption = "Particle - Spin"

End Select

lblVis.Visible = True
tmrVis.Enabled = True

End Sub
