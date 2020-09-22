Attribute VB_Name = "modPlaylist"
Option Explicit

'
' module containing Playlist related sub-routines
'

Public Sub savePlayList(lstPath As ListBox, File As String)
Dim OutStream As TextStream, i As Integer
i = 0
If Trim(File) = "" Then GoTo e
Set OutStream = FSys.CreateTextFile(File, True, False)

'fpl headers
OutStream.WriteLine ("<?fpl version=" & Chr(&H22) & "1.0" & Chr(&H22) & "?>") 'write XML header
OutStream.WriteLine ("<!-- Created on " & Date & "; WARNING !!! This File is Machine Generated. Do NOT Edit. -->") ' write date
OutStream.WriteLine ("<playlist generator=" & Chr(&H22) & "FireAMP"" " & "version=""" & App.Major & "." & App.Minor & "." & App.Revision & Chr(&H22) & ">") ' write main tag

While i < lstPath.ListCount
OutStream.WriteLine "    <path> " & lstPath.List(i) & " </path>" 'write path
OutStream.WriteLine "    <name> " & getFileName(lstPath.List(i)) & " </name>" 'write song name
i = i + 1
Wend

OutStream.WriteLine ("</playlist>")
Set OutStream = Nothing ' destroy object

e:
End Sub

' Loads playlist

Public Sub openPlayList(lstPlaylist As ListView, lstPath As ListBox, File As String)
Dim InStream As TextStream

Dim a As Boolean, str As String
Dim LST As ListItem
Dim i As Integer, ext As String

Let a = False
If Trim(File) = "" Then GoTo e
Set InStream = FSys.OpenTextFile(Trim(File), ForReading, False, TristateFalse)


If Not StrComp(Replace(InStream.Read(5), "<?", " "), "fpl") Then
 Dim e As ErrStruct
 e.errNum = 5
 e.errShortDesc = "This does not appear to be a FireAMP! Playlist"
 e.errLongDesc = "The playlist recently opened did not have the FireAMP! playlist header in it. The File is either corrupt or invalid"
 logError e
Exit Sub
End If

InStream.SkipLine ' Skip header
InStream.SkipLine ' Skip date
InStream.SkipLine ' Skip main tag

lstPath.Clear
lstPlaylist.ListItems.Clear

While InStream.AtEndOfStream = False
str = InStream.readLine
If a = True Then
Set LST = lstPlaylist.ListItems.Add(, , parseString(str, 7, 7)) ' load name
Else
lstPath.AddItem (parseString(str, 7, 7)) ' load path
End If
a = Not a

Wend

Set InStream = Nothing ' destroy object
e:
End Sub

Public Function parseString(Src As String, Start As Integer, Finish As Integer) As String
On Error Resume Next
Dim str As String, str1 As String
Src = Trim(Src)
str = Left(Src, Len(Src) - Start)
str1 = Right(str, Len(str) - Finish)
parseString = str1
End Function

