Attribute VB_Name = "modMedia"
Option Explicit

'
' module containing media related functions
'

Dim FireAMP As QuartzTypeLib.FilgraphManager       ' FireAMP!
Public FireAMP_Pos As QuartzTypeLib.IMediaPosition ' player position
Public FireAMP_Vol As QuartzTypeLib.IBasicAudio    ' player volume
Public FireAMP_VideoWin As QuartzTypeLib.IVideoWindow ' video window
Public currentVolume As Long

Public isPlaying As Boolean
Public curFile As String

Public Function PlayClip(fileToPlay As String, Optional testFile As Boolean = False) As Boolean
On Error GoTo errHandle
 Set FireAMP = New FilgraphManager
 FireAMP.RenderFile (fileToPlay)
 
 
 
If Not testFile Then
 Set FireAMP_Pos = FireAMP
 Set FireAMP_VideoWin = FireAMP
 Set FireAMP_Vol = FireAMP

 frmFireMain.fraVideo.Visible = False
  frmFireMain.fraDisplay.Visible = False
 
 Select Case LCase(getExtension(fileToPlay))
 Case "mpg", "mpeg", "dat", "mov", "wmv":
 If frmFullScreen.Visible Then
 FireAMP_VideoWin.HideCursor True
 refreshVideo frmFullScreen.Frame1
Else
FireAMP_VideoWin.HideCursor False
   refreshVideo frmFireMain.fraVideo
  frmFireMain.fraVideo.Visible = True
End If
 Case Else
  frmFireMain.fraDisplay.Visible = True
 End Select
FireAMP.Run
isPlaying = True
End If

PlayClip = True

Exit Function
errHandle:
If Not testFile Then
 Dim e As ErrStruct
 e.errNum = Err.Number
 e.errShortDesc = "FireAMP internal error: unSupported file type"
 e.errLongDesc = "FireAMP tried to play a file type that is not supported" _
 & " Remove the clip from the playlist if it was added."
 
  logError e
  End If
  
PlayClip = False
Err.Clear
End Function

Public Sub StopClip()
If FireAMP Is Nothing Then Exit Sub
 FireAMP.Stop
 Set FireAMP = Nothing ' release object
 isPlaying = False
End Sub

Public Sub PauseClip()
Static isPaused As Boolean
isPaused = Not isPaused

 If isPaused Then
  FireAMP.Pause
 Else
  FireAMP.Run
 End If
End Sub

Public Sub refreshVideo(objOwner As Frame)

FireAMP_VideoWin.Width = objOwner.Width
FireAMP_VideoWin.Height = objOwner.Height

FireAMP_VideoWin.Left = 0
FireAMP_VideoWin.Top = 0

FireAMP_VideoWin.WindowStyle = CLng(&H6000000)  ' window style: no border
FireAMP_VideoWin.Owner = objOwner.hwnd          ' assign window region

End Sub
