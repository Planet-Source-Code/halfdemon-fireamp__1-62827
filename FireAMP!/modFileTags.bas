Attribute VB_Name = "modFileTags"
Option Explicit

'
' module containing tag reading functions for mp3, midi, wma files
'

'
' mp3
'

' the mp3 id3 version 1 tag structure
Public Type tagMP3ID3V1
  Tag As String * 3           ' 003 byte
  Title As String * 30        ' 033 byte
  Artist As String * 30       ' 063 byte
  Album As String * 30        ' 093 byte
  Year As String * 4          ' 097 byte
  Comment As String * 28      ' 125 byte
  Filler As Byte              ' 126 byte
  Track As Byte               ' 127 byte
  Genre As Byte               ' 128 byte
End Type

' function to get a genre corresponding to a genre number
Public Function getGenre(Genre As Byte)
Dim sName As String
   Select Case Genre
   'A
   Case 34: sName = "Acid"
   Case 74: sName = "Acid Jazz"
   Case 73: sName = "Acid Punk"
   Case 99: sName = "Acoustic"
   Case 40: sName = "Alt.Rock"
   Case 20: sName = "Alternative"
   Case 26: sName = "Ambient"
   Case 145: sName = "Anime"
   Case 90: sName = "Avant Garde"
   
   'B
   Case 116: sName = "Ballad"
   Case 41: sName = "Bass"
   Case 135: sName = "Beat"
   Case 85: sName = "Bebob"
   Case 96: sName = "Big Band"
   Case 138: sName = "Black Metal"
   Case 89: sName = "Blue Grass"
   Case 0: sName = "Blues"
   Case 107: sName = "Booty Bass"
   Case 132: sName = "Brit Pop"
   
   'C
   Case 65: sName = "Cabaret"
   Case 88: sName = "Celtic"
   Case 104: sName = "Chamber Music"
   Case 102: sName = "Chanson"
   Case 97: sName = "Chorus"
   Case 136: sName = "Christian Gangsta Rap"
   Case 61: sName = "Christian Rap"
   Case 141: sName = "Christian Rock"
   Case 1: sName = "Classic Rock"
   Case 32: sName = "Classical"
   Case 112: sName = "Club"
   Case 128: sName = "Club - House"
   Case 57: sName = "Comedy"
   Case 140: sName = "Contemporary Christian"
   Case 2: sName = "Country"
   Case 139: sName = "Crossover"
   Case 58: sName = "Cult"
   
   'D
   Case 3: sName = "Dance"
   Case 125: sName = "Dance Hall"
   Case 50: sName = "Darkwave"
   Case 22: sName = "Death Metal"
   Case 4: sName = "Disco"
   Case 55: sName = "Dream"
   Case 127: sName = "Drum & Bass"
   Case 122: sName = "Drum Solo"
   Case 120: sName = "Duet"
   
   'E
   Case 98: sName = "Easy Listening"
   Case 52: sName = "Electronic"
   Case 48: sName = "Ethnic"
   Case 54: sName = "Eurodance"
   Case 124: sName = "Euro - House"
   Case 25: sName = "Euro - Techno"
   
   'F
   Case 84: sName = "Fast Fusion"
   Case 80: sName = "Folk"
   Case 81: sName = "Folk / Rock"
   Case 115: sName = "Folklore"
   Case 119: sName = "Freestyle"
   Case 5: sName = "Funk"
   Case 30: sName = "Fusion"
   
   'G
   Case 36: sName = "Game"
   Case 59: sName = "Gangsta Rap"
   Case 126: sName = "Goa"
   Case 38: sName = "Gospel"
   Case 49: sName = "Gothic"
   Case 91: sName = "Gothic Rock"
   Case 6: sName = "Grunge"
   
   'H
   Case 79: sName = "Hard Rock"
   Case 129: sName = "Hardcore"
   Case 137: sName = "Heavy Metal"
   Case 7: sName = "Hip Hop"
   Case 35: sName = "House"
   Case 100: sName = "Humour"
   
   'I
   Case 131: sName = "Indie"
   Case 19: sName = "Industrial"
   Case 33: sName = "Instrumental"
   Case 46: sName = "Instrumental Pop"
   Case 47: sName = "Instrumental Rock"
   
   'J
   Case 8: sName = "Jazz"
   Case 29: sName = "Jazz - Funk"
   Case 146: sName = "JPop"
   Case 63: sName = "Jungle"
   
   'L
   Case 86: sName = "Latin"
   Case 71: sName = "Lo - fi"
   
   'M
   Case 45: sName = "Meditative"
   Case 142: sName = "Merengue"
   Case 9: sName = "Metal"
   Case 77: sName = "Musical"
   Case 82: sName = "National Folk"

   'N
   Case 64: sName = "Native American"
   Case 133: sName = "Negerpunk"
   Case 10: sName = "New Age"
   Case 66: sName = "New Wave"
   Case 39: sName = "Noise"
   
   'O
   Case 11: sName = "Oldies"
   Case 103: sName = "Opera"
   Case 12: sName = "Other"
   
   'P
   Case 75: sName = "Polka"
   Case 134: sName = "Polsk Punk"
   Case 13: sName = "Pop"
   Case 62: sName = "Pop / Funk"
   Case 53: sName = "Pop / Folk"
   Case 109: sName = "Pr0n Groove"
   Case 117: sName = "Power Ballad"
   Case 23: sName = "Pranks"
   Case 108: sName = "Primus"
   Case 92: sName = "Progressive Rock"
   Case 67: sName = "Psychedelic"
   Case 93: sName = "Psychedelic Rock"
   Case 43: sName = "Punk"
   Case 121: sName = "Punk Rock"
   
   'R
   Case 14: sName = "R&B"
   Case 15: sName = "Rap"
   Case 68: sName = "Rave"
   Case 16: sName = "Reggae"
   Case 76: sName = "Retro"
   Case 87: sName = "Revival"
   Case 118: sName = "Rhythmic Soul"
   Case 17: sName = "Rock"
   Case 78: sName = "Rock 'n'Roll"
   
   'S
   Case 143: sName = "Salsa"
   Case 114: sName = "Samba"
   Case 110: sName = "Satire"
   Case 69: sName = "Showtunes"
   Case 21: sName = "Ska"
   Case 111: sName = "Slow Jam"
   Case 95: sName = "Slow Rock"
   Case 105: sName = "Sonata"
   Case 42: sName = "Soul"
   Case 37: sName = "Sound Clip"
   Case 24: sName = "Soundtrack"
   Case 56: sName = "Southern Rock"
   Case 44: sName = "Space"
   Case 101: sName = "Speech"
   Case 83: sName = "Swing"
   Case 94: sName = "Symphonic Rock"
   Case 106: sName = "Symphony"
   Case 147: sName = "Synth Pop"

   'T
   Case 113: sName = "Tango"
   Case 18: sName = "Techno"
   Case 51: sName = "Techno - Industrial"
   Case 130: sName = "Terror"
   Case 144: sName = "Thrash Metal"
   Case 60: sName = "Top 40"
   Case 70: sName = "Trailer"
   Case 31: sName = "Trance"
   Case 72: sName = "Tribal"
   Case 27: sName = "Trip Hop"
   
   'V
   Case 28: sName = "Vocal"
   
   End Select
   getGenre = sName
End Function

' function to read tag from an mp3 file
Public Function readMP3Tag(mp3File As String) As tagMP3ID3V1

Dim fNum As Integer
Dim e As ErrStruct ' the error
Dim theTag As tagMP3ID3V1 ' tag to be read
fNum = FreeFile
' fill the tag with info
theTag.Album = "Unknown"
theTag.Artist = "Unknown"
theTag.Comment = "None"
theTag.Genre = 255
theTag.Title = "Unknown"
theTag.Year = "????"
' open mp3 file

If Not FSys.FileExists(mp3File) Then
e.errNum = 1
e.errShortDesc = "FireAMP external error: File not found"
e.errLongDesc = "This error occurs if the specified file was not found " _
& "to read an mp3 id3 v.1 tag from. If you were playing this clip from" _
& " a playlist, check if the file refered to exists or if it has been moved" _
& " or deleted and update the playlist"
' log an error
 logError e
 
Exit Function
End If

On Error GoTo errHandle
' read tag
Open mp3File For Binary As #fNum
   If LOF(fNum) > 128 Then
      Get #fNum, LOF(fNum) - 127, theTag.Tag
         If Not theTag.Tag = "TAG" Then
              readMP3Tag = theTag
         Else
            
            Get #fNum, , theTag.Title
            theTag.Title = toStdString(theTag.Title)
            
            Get #fNum, , theTag.Artist
            theTag.Artist = toStdString(theTag.Artist)
            
            Get #fNum, , theTag.Album
            theTag.Album = toStdString(theTag.Album)
            
            Get #fNum, , theTag.Year
            theTag.Year = toStdString(theTag.Year)
            
            Get #fNum, , theTag.Comment
            theTag.Comment = toStdString(theTag.Comment)
            
            Get #fNum, , theTag.Filler
            Get #fNum, , theTag.Track
            
            
            Get #fNum, , theTag.Genre
                        
         End If
      End If
   
   Close #fNum
   
   With theTag
   .Album = Trim(.Album)
   .Artist = Trim(.Artist)
   .Comment = Trim(.Comment)
  End With
   
   readMP3Tag = theTag
   
Exit Function
errHandle:

 e.errNum = Err.Number
 e.errShortDesc = "FireAMP error: " & Err.Description
 e.errLongDesc = "(Internal error, no help topic associated with this type)"
 logError e

Err.Clear ' must always clear curent error
End Function

Public Sub writeTag(theTag As tagMP3ID3V1, mp3File As String)

Dim fNum As Integer
fNum = FreeFile

   On Error GoTo errHandle
   Open mp3File For Binary Access Read Write Lock Write As #fNum
   If LOF(fNum) > 0 Then
         If LOF(fNum) > 128 Then
            Get #fNum, LOF(fNum) - 127, theTag.Tag
            If Not (StrComp(theTag.Tag, "TAG") = 0) Then
               ' no MP3 tag already, need to extend the file
              
               Seek #fNum, LOF(fNum)
               theTag.Tag = "TAG"
               Put #fNum, , theTag.Tag
            End If
            
            Put #fNum, , theTag.Title
            Put #fNum, , theTag.Artist
            Put #fNum, , theTag.Album
            Put #fNum, , theTag.Year
            Put #fNum, , theTag.Comment
            Put #fNum, , theTag.Filler
            Put #fNum, , theTag.Track
            Put #fNum, , theTag.Genre
      End If
      End If
Close #fNum
Exit Sub
errHandle:
Close #fNum
Dim e As ErrStruct
e.errNum = Err.Number
 e.errShortDesc = "FireAMP error: " & Err.Description
 e.errLongDesc = "FireAMP could not write the tag to the specified mp3 file.The file is either currently playing or" _
 & " is being used by another application. Stop the current clip" _
 & " or try again"
 
logError e

Err.Clear ' must always clear curent error

End Sub

' function to specify allowed characters in a string
' allows only alpha-numeric and special characters
Private Function isAllowed(checkString As String) As Boolean
Dim a As Boolean
a = False
' allowedChar - string contains all allowed characters
Dim allowedChar As String
allowedChar = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz 1234567890-.*" & Chr(&H22) ' chr(H22) is " (double quotes)
If InStr(1, allowedChar, checkString) Then a = True
isAllowed = a
End Function

Public Function readMIDIInfo(MFile As String) As String

Dim tTitle As String * 40
Dim chunk As String * 4
Dim Data As String * 8

Open MFile For Binary Access Read As #1
Get #1, , chunk

 If chunk = "MThd" Then ' midi header chunk
 Get #1, 7, Data
 Get #1, , chunk
  If chunk = "MTrk" Then 'firsr midi track
  Get #1, 19, Data
  Get #1, , tTitle
  MsgBox tTitle
  End If
 End If
 
Close #1


End Function

