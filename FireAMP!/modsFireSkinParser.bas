Attribute VB_Name = "modsFireSkinParser"
'
' module containing skin parser function
'

'not yet perfected

Option Explicit
Public plColor As Long
' colors used in vis.
Public dimColor1 As Long
Public dimColor2 As Long
Public dimColor3 As Long

Public stepColor1 As Long
Public stepColor2 As Long
Public stepColor3 As Long

Public skinName As String
Public skinAuthor As String
Public skinNotes As String

Public Wx As Integer, Wy As Integer, cWx As Integer, cWy As Integer


'the one and only function to layout a skin
Public Sub renderSkin(skinName As String)
'Unload frmFireMain

On Error Resume Next

Dim skinUnPacker As New FireSkinLibrary.FireSkinner

' check if temporary directory exists
If FSys.FolderExists(App.Path & "\Temp") Then
'yes? delete all files
FSys.DeleteFolder App.Path & "\Temp"
FSys.CreateFolder App.Path & "\Temp"

Else
'no? well create one...
FSys.CreateFolder App.Path & "\Temp"
End If

'unpack the skin
skinUnPacker.decodeFireSkin skinName, App.Path & "\temp\"

' we donot need the skin unpacker anymore
Set skinUnPacker = Nothing

'start to parse the files
Dim Fin As TextStream
'locate a .fss file in the temp directory
If FSys.FileExists(App.Path & "\Temp\" & Dir(App.Path & "\Temp\*.fss")) Then
Set Fin = FSys.OpenTextFile(App.Path & "\Temp\" & Dir(App.Path & "\Temp\*.fss"))

Else
Dim e As ErrStruct
e.errNum = 7
e.errShortDesc = "Corrupt skin!"
e.errLongDesc = "The skin specification was not found in the archive"
logError e
Exit Sub
End If


Dim Line$, windowRegion As Long, ln As Integer
Dim parts() As String, rgn() As String
'clear all

Dim Control As Object
For Each Control In frmFireMain
If TypeOf Control Is PictureBox Then Control.Cls
Next



While Not Fin.AtEndOfStream
Line = Fin.ReadLine


' load files
If LCase(Line) = "files" Then
Fin.ReadLine

While Line <> "%>"

Line = Fin.ReadLine
  ln = ln + 1
  parts = Split(Line, ":")
  If UBound(parts) > 0 Then
   parts(1) = Trim(parts(1))
   Select Case parts(0)
   'load main picture
   Case "main":
   Set frmFireMain.picSkin.Picture = LoadPicture(App.Path & "\temp\" & parts(1))
   frmFireMain.Height = frmFireMain.picSkin.Height
   frmFireMain.Width = frmFireMain.picSkin.Width
    windowRegion = MakeRegion(frmFireMain.picSkin)
     SetWindowRgn frmFireMain.hwnd, windowRegion, True
   'load playlist picture
   Case "playlist":
   Set frmFirePL.picSkin.Picture = LoadPicture(App.Path & "\temp\" & parts(1))
   frmFirePL.Width = frmFirePL.picSkin.Width
   frmFirePL.Height = frmFirePL.picSkin.Height
    windowRegion = MakeRegion(frmFirePL.picSkin)
     SetWindowRgn frmFirePL.hwnd, windowRegion, True
   
   'load other pictures
   Case "buttons": frmFireMain.picBtnSrc.Picture = LoadPicture(App.Path & "\temp\" & parts(1))
   Case "controls": frmFireMain.picCtrlSrc.Picture = LoadPicture(App.Path & "\temp\" & parts(1))
   
   Case "seek-bar": frmFireMain.picBarBack.Picture = LoadPicture(App.Path & "\temp\" & parts(1))
   Case "pl-bar": frmFirePL.picBack.Picture = LoadPicture(App.Path & "\temp\" & parts(1))
   Case "seek-bar-front": frmFireMain.picBarFront.Picture = LoadPicture(App.Path & "\temp\" & parts(1))
   Case "pl-bar-front": frmFirePL.picBar.Picture = LoadPicture(App.Path & "\temp\" & parts(1))
   
   Case "media-tracker":
   Unload frmMediaTracker
   frmMediaTracker.picSkin.Picture = LoadPicture(App.Path & "\temp\" & parts(1))
      
   Case "background1": frmFireMain.fraDisplay.Picture = LoadPicture(App.Path & "\temp\" & parts(1))
   Case "background2": frmFireMain.ScopeBuff.Picture = LoadPicture(App.Path & "\temp\" & parts(1))
   
   End Select
  End If
  Wend

End If

'load data
If LCase(Line) = "data" Then

Fin.ReadLine
While Line <> "%>"
 Line = Fin.ReadLine

parts = Split(Line, ":")
  If UBound(parts) > 0 Then
     checkData parts(0), parts(1)
  End If
Wend

End If

' arrange elements
If LCase(Line) = "arrange" Then

Fin.ReadLine
While Line <> "%>"
 Line = Fin.ReadLine
 ln = ln + 1
parts = Split(Line, ":")
  If UBound(parts) > 0 Then
     Arrange parts(0), parts(1)
   
  End If
Wend
End If

' change fonts
If LCase(Line) = "fonts" Then

Fin.ReadLine
While Line <> "%>"
 Line = Fin.ReadLine
parts = Split(Line, ":")
  If UBound(parts) > 0 Then
   changeFont parts(0), parts(1)
  End If
Wend
End If

' change colors
If LCase(Line) = "colors" Then

Fin.ReadLine
While Line <> "%>"
 Line = Fin.ReadLine
parts = Split(Line, ":")
  If UBound(parts) > 0 Then
   changeColor parts(0), parts(1)
  End If
Wend
End If

Wend


With frmFireMain
.picBtnSrc.Refresh
.picCtrlSrc.Refresh

Wx = .picBtnSrc.ScaleWidth / 2
Wy = .picBtnSrc.ScaleHeight / 4

cWx = .picCtrlSrc.ScaleWidth / 2
cWy = .picCtrlSrc.ScaleHeight / 2

BitBlt .picBtn(0).hdc, 0, 0, Wx, Wy, .picBtnSrc.hdc, 0, 0, vbSrcCopy 'play
BitBlt .picBtn(1).hdc, 0, 0, Wx, Wy, .picBtnSrc.hdc, 0, Wx, vbSrcCopy 'stop
BitBlt .picBtn(2).hdc, 0, 0, Wx, Wy, .picBtnSrc.hdc, 0, Wx * 3, vbSrcCopy 'open

BitBlt .picCtrl(0).hdc, 0, 0, cWx, cWy, .picCtrlSrc.hdc, 0, 0, vbSrcCopy
BitBlt .picCtrl(1).hdc, 0, 0, cWx, cWy, .picCtrlSrc.hdc, 0, cWy, vbSrcCopy

.picBtn(0).Refresh
.picBtn(1).Refresh
.picBtn(2).Refresh

.picCtrl(0).Refresh
.picCtrl(1).Refresh

.Frame1.BackColor = GetPixel(.picSkin.hdc, .Frame1.Left, .Frame1.Top)

Dim i

For i = 0 To 2
.picBtn(i).Width = Wx
.picBtn(i).Height = Wy
Next

For i = 0 To 1
.picCtrl(i).Width = cWx
.picCtrl(i).Height = cWy
Next

End With




Set Fin = Nothing
End Sub

Function parseRGB(RGBString As String) As Long

parseRGB = RGB(Val("&h" & Mid(RGBString, 1, 2)), Val("&h" & Mid(RGBString, 3, 2)), Val("&h" & Mid(RGBString, 5, 2)))
End Function

Sub checkData(Line$, Data$)
If Line Like "name" Then
skinName = Data
ElseIf Line Like "author" Then
skinAuthor = Data
ElseIf Line Like "notes" Then
skinNotes = Data
End If
End Sub
Sub Arrange(Line$, Data$)
If Line Like "main-caption" Then
moveObject frmFireMain.lblCaption, Data

ElseIf Line Like "main-seek-bar" Then
moveObject frmFireMain.picBarBack, Data

ElseIf Line Like "main-play-button" Then
moveObject frmFireMain.picBtn(0), Data

ElseIf Line Like "main-stop-button" Then
moveObject frmFireMain.picBtn(1), Data

ElseIf Line Like "main-open-button" Then
moveObject frmFireMain.picBtn(2), Data

ElseIf Line Like "main-close-button" Then
moveObject frmFireMain.picCtrl(0), Data

ElseIf Line Like "main-min-button" Then
moveObject frmFireMain.picCtrl(1), Data

ElseIf Line Like "main-time" Then
moveObject frmFireMain.lblStatus, Data

ElseIf Line Like "main-info" Then
moveObject frmFireMain.Frame1, Data

ElseIf Line Like "pl-caption" Then
moveObject frmFirePL.lblCaption, Data

ElseIf Line Like "pl-list" Then
moveObject frmFirePL.lstPL, Data

ElseIf Line Like "pl-bar" Then
moveObject frmFirePL.picBack, Data

ElseIf Line Like "mt-caption" Then
moveObject frmMediaTracker.lblCaption, Data

ElseIf Line Like "mt-title" Then
moveObject frmMediaTracker.lblTitle, Data

ElseIf Line Like "mt-time" Then
moveObject frmMediaTracker.lblTime, Data

ElseIf Line Like "song-title" Then
moveObject frmFireMain.lblTitle, Data

ElseIf Line Like "song-album" Then
moveObject frmFireMain.lblAlbum, Data

ElseIf Line Like "vis" Then
moveObject frmFireMain.Scope, Data

ElseIf Line Like "vis-caption" Then
moveObject frmFireMain.lblVis, Data

ElseIf Line Like "video" Then
moveObject frmFireMain.fraVideo, Data
moveObject frmFireMain.fraDisplay, Data

End If
End Sub

Sub moveObject(theObject As Object, Data As String)
Dim Regions() As String
Regions = Split(Data, ",")
theObject.Move Regions(0), Regions(1)
theObject.Height = Regions(2)
theObject.Width = Regions(3)
End Sub

Sub changeFont(Line$, Data$)
If Line Like "main-caption" Then
makeFontChange frmFireMain.lblCaption, Data

ElseIf Line Like "main-title" Then
makeFontChange frmFireMain.lblTitle, Data

ElseIf Line Like "main-album" Then
makeFontChange frmFireMain.lblAlbum, Data

ElseIf Line Like "main-time" Then
makeFontChange frmFireMain.lblStatus, Data

ElseIf Line Like "main-info" Then
makeFontChange frmFireMain.lblInfo, Data

ElseIf Line Like "pl-caption" Then
makeFontChange frmFirePL.lblCaption, Data

ElseIf Line Like "pl-list" Then
makeFontChange frmFirePL.lstPL, Data

ElseIf Line Like "mt-caption" Then
makeFontChange frmMediaTracker.lblCaption, Data

ElseIf Line Like "mt-title" Then
makeFontChange frmMediaTracker.lblTitle, Data

ElseIf Line Like "mt-time" Then
makeFontChange frmMediaTracker.lblTime, Data

End If
End Sub

Sub makeFontChange(theObject As Object, Data As String)
Dim Regions() As String
Regions = Split(Data, ",")

theObject.Font.Name = Regions(0)
theObject.Font.Size = Val(Regions(1))

Select Case LCase(Regions(2))
 Case "b"
 theObject.Font.Bold = True
 Case "i"
  theObject.Font.Italic = True
 Case "u"
  theObject.Font.Underline = True
 Case "s"
 theObject.Font.Strikethrough = True
  Case "n"
 theObject.Font.Bold = False
 theObject.Font.Italic = False
 theObject.Font.Underline = False
 theObject.Font.Strikethrough = False
 End Select
 
End Sub

Sub changeColor(Line$, Data$)
If Line Like "main-caption" Then
makeColorChange frmFireMain.lblCaption, Data

ElseIf Line Like "main-title" Then
makeColorChange frmFireMain.lblTitle, Data

ElseIf Line Like "main-album" Then
makeColorChange frmFireMain.lblAlbum, Data

ElseIf Line Like "main-time" Then
makeColorChange frmFireMain.lblStatus, Data

ElseIf Line Like "main-info" Then
makeColorChange frmFireMain.lblInfo, Data

ElseIf Line Like "pl-caption" Then
makeColorChange frmFirePL.lblCaption, Data

ElseIf Line Like "pl-list" Then
makeColorChange frmFirePL.lstPL, Data
plColor = parseRGB(Data)

ElseIf Line Like "pl-back-list" Then
frmFirePL.lstPL.BackColor = parseRGB(Data)

ElseIf Line Like "mt-caption" Then
makeColorChange frmMediaTracker.lblCaption, Data

ElseIf Line Like "mt-title" Then
makeColorChange frmMediaTracker.lblTitle, Data

ElseIf Line Like "mt-time" Then
makeColorChange frmMediaTracker.lblTime, Data

ElseIf Line Like "vis-box-fore" Then
makeColorChange frmFireMain.ScopeBuff, Data

ElseIf Line Like "vis-box?" Then
Select Case Right(Line, 1)
Case "0"
frmFireMain.fraDisplay.BackColor = parseRGB(Data)
frmFireMain.Scope.BackColor = parseRGB(Data)
frmFireMain.ScopeBuff.BackColor = parseRGB(Data)

Case "1"
dimColor1 = parseRGB(Data)
Case "2"
dimColor2 = parseRGB(Data)
Case "3"
dimColor3 = parseRGB(Data)
End Select

ElseIf Line Like "vis-box-step?" Then
Select Case Right(Line, 1)
Case "1"
 stepColor1 = Val(Data)
Case "2"
 stepColor2 = Val(Data)
Case "3"
 stepColor3 = Val(Data)
End Select
End If
End Sub

Sub makeColorChange(theObject As Object, Data As String)
theObject.ForeColor = parseRGB(Data)
End Sub
