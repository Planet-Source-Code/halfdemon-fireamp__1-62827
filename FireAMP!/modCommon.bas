Attribute VB_Name = "modCommon"
Option Explicit

'
' module containing commonly used functions and subroutines
'

Global FSys As New FileSystemObject 'global FileSystemObject
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
        (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
        lParam As Any) As Long

Const CB_FINDSTRING = &H14C
Const CB_FINDSTRINGEXACT = &H158
Const LB_FINDSTRING = &H18F
Const LB_FINDSTRINGEXACT = &H1A2



' type to encapsulate errors
Public Type ErrStruct
  errNum As Long
  errShortDesc As String
  errLongDesc As String
End Type
Public fraVideoWin As Frame
Public Abt As Boolean

Public Type FireAMPoptions
' general
showMediaTracker As Byte
enableVisualizations As Byte

'start up
showSplashScreen As Byte
loadDefaultSkin As Byte
checkAssociationsAtStartUp As Byte

' file types
MIDI As Byte
WAV As Byte
MP3 As Byte
MPG As Byte
WMA As Byte

End Type
Public theOptions As FireAMPoptions

' subroutine to log errors
Public Sub logError(theError As ErrStruct)
 
Dim FOut As TextStream
 With frmFireTrap
  .lblError = theError.errShortDesc
  .lblReason = theError.errLongDesc
  .lblNum = "Error #" & theError.errNum
 End With
frmFireTrap.Show vbModal

' log error to file
 Set FOut = FSys.OpenTextFile(App.Path & "\FireAMP.Errors.Log", ForAppending, True)
 If FSys.GetFile(App.Path & "\FireAMP.Errors.Log").Size > 10& * 1024& Then ' greater than 10kb
  FOut.Close
  Kill App.Path & "\FireAMP.Errors.Log"
  Set FOut = FSys.OpenTextFile(App.Path & "\FireAMP.Errors.Log", ForAppending, True)
 End If
FOut.WriteLine "FireAMP error #" & theError.errNum
FOut.WriteLine "Error occured on: " & Now
FOut.WriteLine "Short Desc: " & theError.errShortDesc
FOut.WriteLine "Long Desc: " & theError.errLongDesc
FOut.WriteLine String(40, "-")
FOut.Close

End Sub

Private Function isAllowedChar(testStr As String) As Boolean
 isAllowedChar = InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZ abcdefghijklmnopqrstuvwxyz1234567890(){}[]!;:'"",.$%*+#|\/", testStr)
End Function

' function to remove unwnated chars
Public Function toStdString(theString As String) As String
Dim retStr As String, i As Integer, j As Integer
retStr = Space(Len(theString)) ' fill up return string
Let j = 1
For i = 1 To Len(theString)

 If isAllowedChar(Mid(theString, i, 1)) Then
  ' mid is more faster and efficient than '&'
  Mid(retStr, j, 1) = Mid(theString, i, 1)
  j = j + 1
 End If
Next i

toStdString = Trim(retStr)
 
End Function

' retrives the file name with extension
Public Function getFileName(FilePath As String) As String
 getFileName = Right(FilePath, Len(FilePath) - InStrRev(FilePath, "\"))
End Function

' gets the file title

Public Function getFileTitle(FilePath As String)
 getFileTitle = LCase(Right(FilePath, Len(FilePath) - InStrRev(FilePath, ".")))
End Function

Public Function getExtension(FilePath As String) As String
 getExtension = Replace(LCase(Right(Trim(FilePath), 4)), ".", "")
End Function

' function to convert seconds to HH:MM:SS format
Public Function convertToStdTime(ByVal iSeconds As Long) As String

'Format input value to "00:00:00"
Dim HH As Long                   'Hours
Dim MM As Long                   'Minutes
Dim SS As Long                   'Seconds
Dim Tmp As String                'Temporary value

 'Old values time is made of
 HH = iSeconds \ 3600
 MM = iSeconds \ 60 Mod 60
 SS = iSeconds Mod 60
 
 'If there is hour
 If HH > 0 Then Tmp = Format$(HH, "00:")
 'Format input
 convertToStdTime = Tmp & Format$(MM, "00:") & Format$(SS, "00")
End Function

Public Function getBarPosition(picBar As PictureBox, picBarBack As PictureBox, iMax As Integer) As Single
getBarPosition = (picBar.Left * iMax) / picBarBack.ScaleWidth
End Function

Sub Main()

If FSys.FileExists(App.Path & "\FireAMP.Options") Then
Open App.Path & "\FireAMP.Options" For Binary Access Read As 1
Get #1, , theOptions
Close #1
End If


If theOptions.showSplashScreen = 1 Then
Unload frmFireMain
frmSplash.Show
Else
frmFireMain.Show
End If

'registerType ".fpl", "FireAMP Playlist", "FireAMP"
'If theOptions.MP3 Then registerType ".mp3", "FireAMP MP3 Audio", "MP3"
'
'If theOptions.MIDI Then
'registerType ".mid", "FireAMP Sequence(MID)", "MIDI"
'registerType ".rmi", "FireAMP Sequence(RMI)", "RMI"
'End If
'
'If theOptions.MPG Then
'registerType ".mpg", "FireAMP Video(MPG)", "MPG"
'registerType ".mpe", "FireAMP Video(MPE)", "MPG"
'registerType ".mpeg", "FireAMP Video(MPEG)", "MPEG"
'End If
'
'If theOptions.WAV Then registerType ".wav", "FireAMP Audio(WAVE)", "WAV"
'
'If theOptions.WMA Then
'registerType ".wma", "Windows Media Audio", "WMA"
'registerType ".wmv", "Windows Media Video", "WMV"
'End If
'
If App.PrevInstance Then End

If Command$ <> "" Then
Select Case getExtension(Command$)
Case "fpl"
 openPlayList frmFirePL.lstPL, frmFirePL.lstPaths, Command$
Case "mp3", "mp2", "mp1", "wma", "wmv", "mid", "rmi", "mpg", "mpeg", "mpe"

 curFile = Replace(Command, """", "")
 frmFirePL.lstPL.ListItems.Add , , getFileName(Command)
 frmFirePL.lstPaths.AddItem Command$
 frmFireMain.picBtn_Click (0)
 
End Select
End If

End Sub
