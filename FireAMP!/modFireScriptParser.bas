Attribute VB_Name = "modFireScriptParser"
' not yet implemented
Option Explicit

'
' module containing the FireScriptParser 1.0
'

'
' scripting language for FireAMP's skin
'

'
' four dimensions: left, top, height, width
'

Public Type fourDimensions
 dLeft As Long
 dTop As Long
 dHeight As Long
 dWidth As Long
End Type

'
' button size: x, y, alignment (width:w / height:h)
'
Public Type ButtonSize
  X As Integer
  Y As Integer
  Align As String * 1
End Type

'
' two dimensions: left, top
'

Public Type twoDimensions
  dTop As Long
  dLeft As Long
End Type

'
' start up types
'

Public Enum StartUpTypes
 StartUpStatic = 0
 StartUpDynamic = 1
End Enum

'
' control properties: encapsulates various control attributes
'

Public Type ControlProperties
 Align As fourDimensions
 fontFace As String
 fontSize As Integer
 fontAttributes As String * 1
 foreColor As Long
 backColor As Long
End Type

'
' type to encapsulate skin properties
'

'
' the core of FireScript
'

Public Type FireSkinProperties

 SourceFile As String                  ' Skin picture File
 
 Name As String                        ' Skin name
 Author As String                      ' skin author
 creationDate As String                ' skin creation date
 Comment As String                     ' comment
 
 MainSkinDimensions As fourDimensions  ' main skin doimensions
 PlaylistDimensions  As fourDimensions ' playlist dimensions
 auxPLDimensions As fourDimensions     ' auxillary playlist dimensions
 TrackDimensions As fourDimensions     ' Media Tracker dimensions
 
 ControlButtonSize1 As ButtonSize      ' Size of control buttons (play/pause,stop,open)
 ControlButtonSize2 As ButtonSize      ' Size of exit,minimize buttons
 
 PlayButton As twoDimensions           ' Play Button
 StopButton As twoDimensions           ' Stop Button
 PauseButton As twoDimensions          ' Pause Button
 OpenButton As twoDimensions           ' Open Button
 
 ExitButton As twoDimensions           ' Exit Button
 MinimizeButton As twoDimensions       ' Minimize button
 
 NumberSize As ButtonSize              ' size of digits in image
 AniSize As ButtonSize                 ' size of animation
 FormatSize As ButtonSize              ' size of file format info.
 
 MainBarText As String * 1             ' text on main track bar
 TrackBarText As String * 1            ' text on media tracker bar
 
 Numbers As twoDimensions              ' coordinates of digits
 Ani As twoDimensions                  ' coordinates of animation
 Formats As twoDimensions              ' coordinates of file format info
 MainBar As fourDimensions             ' coordinates of main track bar
 TrackBar As fourDimensions            ' coordinates of media tracker track bar
 
 StartUP As StartUpTypes               ' type of start up
 hasMainPlaylist As Boolean            ' has main playlist ?
 
 
 ' controls, names speak of themselves
 MainTitle As ControlProperties
 MainSongTitle As ControlProperties
 MainSongArtist As ControlProperties
 MainSongAlbum As ControlProperties
 
 MainPlayButton As ControlProperties
 MainStopButton As ControlProperties
 MainOpenButton As ControlProperties
 MainSeekBar As ControlProperties
 
 MainExitButton As ControlProperties
 MainMinimizeButton As ControlProperties
 
 MainStatus As ControlProperties
 
 TrackTitle As ControlProperties
 TrackSongTitle As ControlProperties
 TrackTime As ControlProperties
 TrackSeekBar As ControlProperties
 
 AnimateBox As ControlProperties
 Digit_1 As ControlProperties
 Digit_2 As ControlProperties
 Digit_3 As ControlProperties
 Digit_4 As ControlProperties
 
 Colon As ControlProperties
 FileFormatInfo As ControlProperties
 TrackInfo As ControlProperties
 
 
 PlaylistList As ControlProperties
 AuxList As ControlProperties

 
End Type

Dim curSkin As String ' reserved
'===============================================================================

'
' function to parse a FireScript file and retrun a FireSkinProperties object
' containing all the information in the script
'
Public Function parseSkin(theSkin As String) As FireSkinProperties
curSkin = theSkin

Dim Fin As TextStream
Set Fin = FSys.OpenTextFile(theSkin, ForReading)

' useful vars
Dim readLine$
Dim rgn() As String, parts() As String
Dim i  As Integer
Dim curSkinProperties As FireSkinProperties
Dim line As Long
Dim Temp As String

Dim e As ErrStruct


' begin loop
While Not Fin.AtEndOfStream

line = line + 1 'lines read so far
   readLine = Fin.readLine
   
   
   If readLine Like "!*" Then
    GoTo commentReached ' reached a comment
    
   ElseIf InStr(1, readLine, "#region") Then ' entered the region block
       
       rgn = Split(readLine, " ")
       rgn(1) = LCase(rgn(1))
       
        Select Case rgn(1)
        
        Case "documentation"
          'parse documentation
            While LCase(readLine) <> "#end region"
ContinueDocumentation:

             readLine = Fin.readLine
             line = line + 1
             
             If Trim(readLine) = "" Then GoTo ContinueDocumentation
             rgn = Split(readLine, ":")
                          
             ' documentation types
              
                  Select Case Trim(LCase(rgn(0)))
                  
                  Case "@name"
                   curSkinProperties.Name = rgn(1)
                  Case "@author"
                   curSkinProperties.Author = rgn(1)
                  Case "@date"
                   curSkinProperties.creationDate = rgn(1)
                  Case "@comment"
                  rgn(1) = Replace(rgn(1), "`", " ")
                   curSkinProperties.Comment = rgn(1)
                  Case "#end region"
                  Case ""

                  Case Else
                  ' documentation error
                  e.errNum = 4
                    e.errShortDesc = "Script Error, Undefined documentation region at " & line
                    e.errLongDesc = "The current script file contained an invalid documentation region: " & rgn(1) & "" _
                        & " Edit the script to fix this problem"
                   logError e
                    End Select

            Wend
            ' end of documentation region
'=======================================================================================================
        Case "skin"
          'parse skin
          
          While LCase(readLine) <> "#end region"
          
ContinueSkin:
           readLine = Fin.readLine
           line = line + 1
           If Trim(readLine) = "" Then GoTo ContinueSkin
           rgn = Split(readLine, "@")
           
            Select Case Trim(LCase(rgn(0)))
            ' skin regions
             Case "src"
              curSkinProperties.SourceFile = Mid(theSkin, 1, Len(theSkin) - Len(getFileName(theSkin))) & Trim(rgn(1))
             Case "main"
               parts = Split(rgn(1), ",")
               curSkinProperties.MainSkinDimensions.dLeft = parts(0)
               curSkinProperties.MainSkinDimensions.dTop = parts(1)
               curSkinProperties.MainSkinDimensions.dHeight = parts(2)
               curSkinProperties.MainSkinDimensions.dWidth = parts(3)
            Case "playlist":
               parts = Split(rgn(1), ",")
               curSkinProperties.PlaylistDimensions.dLeft = parts(0)
               curSkinProperties.PlaylistDimensions.dTop = parts(1)
               curSkinProperties.PlaylistDimensions.dHeight = parts(2)
               curSkinProperties.PlaylistDimensions.dWidth = parts(3)
            Case "aux"
            parts = Split(rgn(1), ",")
               curSkinProperties.auxPLDimensions.dLeft = parts(0)
               curSkinProperties.auxPLDimensions.dTop = parts(1)
               curSkinProperties.auxPLDimensions.dHeight = parts(2)
               curSkinProperties.auxPLDimensions.dWidth = parts(3)
            Case "track"
            parts = Split(rgn(1), ",")
               curSkinProperties.TrackDimensions.dLeft = parts(0)
               curSkinProperties.TrackDimensions.dTop = parts(1)
               curSkinProperties.TrackDimensions.dHeight = parts(2)
               curSkinProperties.TrackDimensions.dWidth = parts(3)
            Case "#end region"
            Case Else
            
            ' skin error
            e.errNum = 5
            e.errShortDesc = "Script Error: Undefined skin region at line " & line
            e.errLongDesc = "The FireAMP script parser encountered a parse error" _
            & ". An invalid skin region was found: " & rgn(0) & ". Edit the script" _
            & " to fix this error"
            logError e
            End Select
          
          Wend
          
          ' end of skin region
'=======================================================================================================
        Case "buttons"
           'parse buttons
           
           While readLine <> "#end region"
           
ContinueButtons:
                   readLine = Fin.readLine
                     line = line + 1
             If Trim(readLine) = "" Then GoTo ContinueButtons
           
             
             If InStr(1, readLine, "?main-buttonsize") Then ' attribute
                               
                
               rgn = Split(readLine, "=")
               parts = Split(rgn(1), ",")
               
               curSkinProperties.ControlButtonSize1.X = parts(0)
               curSkinProperties.ControlButtonSize1.Y = parts(1)
               curSkinProperties.ControlButtonSize1.Align = parts(2)
               ElseIf InStr(1, readLine, "?ctrl-buttonsize") Then
               rgn = Split(readLine, "=")
               parts = Split(rgn(1), ",")
               
               curSkinProperties.ControlButtonSize2.X = parts(0)
               curSkinProperties.ControlButtonSize2.Y = parts(1)
               curSkinProperties.ControlButtonSize2.Align = parts(2)
               ElseIf InStr(1, readLine, "@") Then
               
               If readLine <> "#end region" Then
               
               rgn = Split(readLine, "@")
               parts = Split(rgn(1), ",")
               Select Case LCase(Trim(rgn(0)))
                Case "play"
                  curSkinProperties.PlayButton.dTop = Val(parts(0))
                  curSkinProperties.PlayButton.dLeft = Val(parts(1))
                Case "stop"
                    curSkinProperties.StopButton.dTop = Val(parts(0))
                    curSkinProperties.StopButton.dLeft = Val(parts(1))
                Case "pause"
                    curSkinProperties.PauseButton.dTop = Val(parts(0))
                    curSkinProperties.PauseButton.dLeft = Val(parts(1))
                Case "open"
                    curSkinProperties.OpenButton.dTop = Val(parts(0))
                    curSkinProperties.OpenButton.dLeft = Val(parts(1))
                Case "exit"
                    curSkinProperties.ExitButton.dTop = Val(parts(0))
                    curSkinProperties.ExitButton.dLeft = Val(parts(1))
                Case "minimize"
                    curSkinProperties.MinimizeButton.dTop = Val(parts(0))
                    curSkinProperties.MinimizeButton.dLeft = Val(parts(1))
                Case Else
                 'INS: invalid button error
                End Select
               End If
            Else
            ' INS: invalid button region
        End If
           Wend
           ' end of buttons region
'=======================================================================================================
        Case "controls"
           'parse controls
           While readLine <> "#end region"
                      
ContinueControls:
                      
            readLine = Fin.readLine
            line = line + 1
            If Trim(readLine) = "" Then GoTo ContinueControls
            
            If InStr(1, readLine, "?") Then
             rgn = Split(readLine, "=")
              parts = Split(rgn(1), ",")
             Select Case Trim(LCase(rgn(0)))
             
                Case "?numbers"
                
                 curSkinProperties.NumberSize.X = Val(parts(0))
                 curSkinProperties.NumberSize.Y = Val(parts(1))
                 curSkinProperties.NumberSize.Align = parts(2)
               Case "?ani"
                curSkinProperties.AniSize.X = Val(parts(0))
                curSkinProperties.AniSize.Y = Val(parts(1))
                curSkinProperties.AniSize.Align = parts(2)
               Case "?formats"
                curSkinProperties.FormatSize.X = Val(parts(0))
                curSkinProperties.FormatSize.Y = Val(parts(1))
                curSkinProperties.FormatSize.Align = parts(2)
               Case "?bar-text"
                curSkinProperties.MainBarText = Replace(parts(0), "`", "")
               Case "?track-bar-text"
                curSkinProperties.TrackBarText = Replace(parts(0), "`", "")
              Case Else
              ' INS: invalid attribute error
             End Select
            ElseIf InStr(1, readLine, "@") Then
            rgn = Split(readLine, "@")
            parts = Split(rgn(1), ",")
            
            Select Case LCase(Trim(rgn(0)))
             Case "numbers"
              curSkinProperties.Numbers.dTop = Val(parts(0))
              curSkinProperties.Numbers.dLeft = Val(parts(1))
            Case "ani"
             curSkinProperties.Ani.dTop = Val(parts(0))
             curSkinProperties.Ani.dLeft = Val(parts(1))
            Case "formats"
             curSkinProperties.Formats.dTop = Val(parts(0))
             curSkinProperties.Formats.dLeft = Val(parts(1))
            Case "bar"
             curSkinProperties.MainBar.dLeft = Val(parts(0))
             curSkinProperties.MainBar.dTop = Val(parts(1))
             curSkinProperties.MainBar.dHeight = Val(parts(2))
             curSkinProperties.MainBar.dLeft = Val(parts(3))
            Case "track-bar"
             curSkinProperties.TrackBar.dLeft = Val(parts(0))
             curSkinProperties.TrackBar.dTop = Val(parts(1))
             curSkinProperties.TrackBar.dHeight = Val(parts(2))
             curSkinProperties.TrackBar.dLeft = Val(parts(3))
            Case Else
             'INS: invalid control region
            End Select
            
            Else
              'INS: invalid control attribute error
            End If
                      
           Wend
           ' end of controls region
           
'=======================================================================================================
        Case "general"
            'parse general
            
          While readLine <> "#end region"
          
ContinueGeneral:
            readLine = Fin.readLine
            line = line + 1
            If Trim(readLine) = "" Then GoTo ContinueGeneral
            
              If InStr(1, readLine, "?") Then
              
               rgn = Split(readLine, "=")
               parts = Split(rgn(1), ",")
               Select Case Trim(LCase(rgn(0)))
               Case "?startup":
                 Select Case Trim(LCase(rgn(1)))
                 Case "static"
                 curSkinProperties.StartUP = StartUpStatic
                 Case "dynamic"
                 curSkinProperties.StartUP = StartUpDynamic
                 Case Else
                 'INS: Invalid Startup Type
                 End Select
               Case "?playlist":
               curSkinProperties.hasMainPlaylist = CBool(parts(0))
               Case Else
               'INS: Invalid Attribute in general rgn
            End Select
           
              
              Else
              'INS: invalid attribute,expected '?'
             End If
          Wend
          ' end of general region
'=======================================================================================================
        Case "dynamic"
            'parse dynamic
            'not yet implemented
'=======================================================================================================
        Case "arrange"
            'parse arrange
            
            While readLine <> "#end region"
ContinueArrange:
             readLine = Fin.readLine
             line = line + 1
             If Trim(readLine) = "" Then GoTo ContinueArrange
             
              If readLine <> "#end region" Then
              rgn = Split(readLine, "@")
              parts = Split(rgn(1), ",")
              
                
              Select Case Trim(LCase(rgn(0)))
              Case "main-title":
               curSkinProperties.MainTitle.Align.dLeft = parts(0)
               curSkinProperties.MainTitle.Align.dTop = parts(1)
               curSkinProperties.MainTitle.Align.dHeight = parts(2)
               curSkinProperties.MainTitle.Align.dWidth = parts(3)
              Case "main-song-title":
               curSkinProperties.MainSongTitle.Align.dLeft = parts(0)
               curSkinProperties.MainSongTitle.Align.dTop = parts(1)
               curSkinProperties.MainSongTitle.Align.dHeight = parts(2)
               curSkinProperties.MainSongTitle.Align.dWidth = parts(3)

              Case "main-song-artist":
               curSkinProperties.MainSongArtist.Align.dLeft = parts(0)
               curSkinProperties.MainSongArtist.Align.dTop = parts(1)
               curSkinProperties.MainSongArtist.Align.dHeight = parts(2)
               curSkinProperties.MainSongArtist.Align.dWidth = parts(3)

              Case "main-song-album":
               curSkinProperties.MainSongAlbum.Align.dLeft = parts(0)
               curSkinProperties.MainSongAlbum.Align.dTop = parts(1)
               curSkinProperties.MainSongAlbum.Align.dHeight = parts(2)
               curSkinProperties.MainSongAlbum.Align.dWidth = parts(3)

              Case "main-play-button":
               curSkinProperties.MainPlayButton.Align.dLeft = parts(0)
               curSkinProperties.MainPlayButton.Align.dTop = parts(1)
               
              Case "main-stop-button":
               curSkinProperties.MainStopButton.Align.dLeft = parts(0)
               curSkinProperties.MainStopButton.Align.dTop = parts(1)
               
              Case "main-open-button":
               curSkinProperties.MainOpenButton.Align.dLeft = parts(0)
               curSkinProperties.MainOpenButton.Align.dTop = parts(1)
               
              Case "main-seek-bar":
               curSkinProperties.MainSeekBar.Align.dLeft = parts(0)
               curSkinProperties.MainSeekBar.Align.dTop = parts(1)
               
              Case "main-exit-button":
               curSkinProperties.MainExitButton.Align.dLeft = parts(0)
               curSkinProperties.MainExitButton.Align.dTop = parts(1)
               
              Case "main-minimize-button":
               curSkinProperties.MainMinimizeButton.Align.dLeft = parts(0)
               curSkinProperties.MainMinimizeButton.Align.dTop = parts(1)
               
              Case "main-status":
               curSkinProperties.MainStatus.Align.dLeft = parts(0)
               curSkinProperties.MainStatus.Align.dTop = parts(1)
               curSkinProperties.MainStatus.Align.dHeight = parts(2)
               curSkinProperties.MainStatus.Align.dWidth = parts(3)

              Case "track-title":
               curSkinProperties.TrackTitle.Align.dLeft = parts(0)
               curSkinProperties.TrackTitle.Align.dTop = parts(1)
               curSkinProperties.TrackTitle.Align.dHeight = parts(2)
               curSkinProperties.TrackTitle.Align.dWidth = parts(3)

              Case "track-song-title":
               curSkinProperties.TrackSongTitle.Align.dLeft = parts(0)
               curSkinProperties.TrackSongTitle.Align.dTop = parts(1)
               curSkinProperties.TrackSongTitle.Align.dHeight = parts(2)
               curSkinProperties.TrackSongTitle.Align.dWidth = parts(3)

              Case "track-time":
               curSkinProperties.TrackTime.Align.dLeft = parts(0)
               curSkinProperties.TrackTime.Align.dTop = parts(1)
               curSkinProperties.TrackTime.Align.dHeight = parts(2)
               curSkinProperties.TrackTime.Align.dWidth = parts(3)

              Case "track-seek-bar":
               curSkinProperties.TrackSeekBar.Align.dLeft = parts(0)
               curSkinProperties.TrackSeekBar.Align.dTop = parts(1)
               
              Case "animate":
               curSkinProperties.AnimateBox.Align.dLeft = parts(0)
               curSkinProperties.AnimateBox.Align.dTop = parts(1)
               
              Case "number1":
               curSkinProperties.Digit_1.Align.dLeft = parts(0)
               curSkinProperties.Digit_1.Align.dTop = parts(1)
               
              Case "number2":
               curSkinProperties.Digit_2.Align.dLeft = parts(0)
               curSkinProperties.Digit_2.Align.dTop = parts(1)
               
              Case "number3":
               curSkinProperties.Digit_3.Align.dLeft = parts(0)
               curSkinProperties.Digit_3.Align.dTop = parts(1)
               
              Case "number4":
               curSkinProperties.Digit_4.Align.dLeft = parts(0)
               curSkinProperties.Digit_4.Align.dTop = parts(1)
               
              Case "colon":
               curSkinProperties.Colon.Align.dLeft = parts(0)
               curSkinProperties.Colon.Align.dTop = parts(1)
               
              Case "format":
               curSkinProperties.FileFormatInfo.Align.dLeft = parts(0)
               curSkinProperties.FileFormatInfo.Align.dTop = parts(1)
               
              Case "track":
               curSkinProperties.TrackInfo.Align.dLeft = parts(0)
               curSkinProperties.TrackInfo.Align.dTop = parts(1)
               
              Case "playlist-list":
               curSkinProperties.PlaylistList.Align.dLeft = parts(0)
               curSkinProperties.PlaylistList.Align.dTop = parts(1)
               curSkinProperties.PlaylistList.Align.dHeight = parts(2)
               curSkinProperties.PlaylistList.Align.dWidth = parts(3)

              Case "aux-list":
               curSkinProperties.AuxList.Align.dLeft = parts(0)
               curSkinProperties.AuxList.Align.dTop = parts(1)
               curSkinProperties.AuxList.Align.dHeight = parts(2)
               curSkinProperties.AuxList.Align.dWidth = parts(3)

              Case Else
              'INS: invalid arrange region
              End Select
             End If
            Wend
            
'=======================================================================================================
        Case "fonts"
            'parse fonts
         While readLine <> "#end region"
ContinueFonts:
          readLine = Fin.readLine
          line = line + 1
          If Trim(readLine) = "" Then GoTo ContinueFonts
          If readLine <> "#end region" Then
          rgn = Split(readLine, "=")
          parts = Split(rgn(1), ",")
          
            Select Case LCase(Trim(rgn(0)))
            Case "main-title":
             curSkinProperties.MainTitle.fontFace = parts(0)
             curSkinProperties.MainTitle.fontSize = Val(parts(1))
             curSkinProperties.MainTitle.fontAttributes = parts(2)
            Case "main-status":
             curSkinProperties.MainStatus.fontFace = parts(0)
             curSkinProperties.MainStatus.fontSize = Val(parts(1))
             curSkinProperties.MainStatus.fontAttributes = parts(2)
            Case "main-song-title":
             curSkinProperties.MainSongTitle.fontFace = parts(0)
             curSkinProperties.MainSongTitle.fontSize = Val(parts(1))
             curSkinProperties.MainSongTitle.fontAttributes = parts(2)

            Case "main-song-artist":
             curSkinProperties.MainSongArtist.fontFace = parts(0)
             curSkinProperties.MainSongArtist.fontSize = Val(parts(1))
             curSkinProperties.MainSongArtist.fontAttributes = parts(2)

            Case "main-song-album":
             curSkinProperties.MainSongAlbum.fontFace = parts(0)
             curSkinProperties.MainSongAlbum.fontSize = Val(parts(1))
             curSkinProperties.MainSongAlbum.fontAttributes = parts(2)

            Case "main-track":
             curSkinProperties.MainSeekBar.fontFace = parts(0)
             curSkinProperties.MainSeekBar.fontSize = Val(parts(1))
             curSkinProperties.MainSeekBar.fontAttributes = parts(2)

            Case "track-title":
             curSkinProperties.TrackSeekBar.fontFace = parts(0)
             curSkinProperties.TrackSeekBar.fontSize = Val(parts(1))
             curSkinProperties.TrackSeekBar.fontAttributes = parts(2)

            Case "track-song-title":
             curSkinProperties.TrackSongTitle.fontFace = parts(0)
             curSkinProperties.TrackSongTitle.fontSize = Val(parts(1))
             curSkinProperties.TrackSongTitle.fontAttributes = parts(2)

            Case "track-time":
             curSkinProperties.TrackTime.fontFace = parts(0)
             curSkinProperties.TrackTime.fontSize = Val(parts(1))
             curSkinProperties.TrackTime.fontAttributes = parts(2)

            Case "playlist-list":
             curSkinProperties.PlaylistList.fontFace = parts(0)
             curSkinProperties.PlaylistList.fontSize = Val(parts(1))
             curSkinProperties.PlaylistList.fontAttributes = parts(2)

            Case "aux-list":
             curSkinProperties.AuxList.fontFace = parts(0)
             curSkinProperties.AuxList.fontSize = Val(parts(1))
             curSkinProperties.AuxList.fontAttributes = parts(2)

            Case Else
            'INS: error- invalid region
            End Select
          End If
          
         Wend
         ' end of fonts region
            
'=======================================================================================================
        Case "colors"
            'parse colors
            
            While readLine <> "#end region"
ContinueColors:
              readLine = Fin.readLine
              line = line + 1
              If Trim(readLine) = "" Then GoTo ContinueColors
              If readLine <> "#end region" Then
              rgn = Split(readLine, "=")
               Select Case LCase(Trim(rgn(0)))
               Case "main-title":
                curSkinProperties.MainTitle.foreColor = Val(rgn(1))
               Case "main-song-title":
                curSkinProperties.MainTitle.foreColor = Val(rgn(1))
               Case "main-song-artist":
                curSkinProperties.MainSongArtist.foreColor = Val(rgn(1))
               Case "main-song-album":
                curSkinProperties.MainSongAlbum.foreColor = Val(rgn(1))
               Case "main-track":
                curSkinProperties.MainSeekBar.foreColor = Val(rgn(1))
               Case "main-status":
                curSkinProperties.MainStatus.foreColor = Val(rgn(1))
               Case "track-title":
                curSkinProperties.TrackTitle.foreColor = Val(rgn(1))
               Case "track-song":
                curSkinProperties.TrackSongTitle.foreColor = Val(rgn(1))
               Case "track-time":
                curSkinProperties.TrackTime.foreColor = Val(rgn(1))
               Case "track-seek-bar":
                curSkinProperties.TrackSeekBar.foreColor = Val(rgn(1))
               Case "playlist-list-fore":
                curSkinProperties.PlaylistList.foreColor = Val(rgn(1))
               Case "playlist-list-back":
                curSkinProperties.PlaylistList.backColor = Val(rgn(1))
               Case "aux-list-fore":
                curSkinProperties.AuxList.foreColor = Val(rgn(1))
               Case "aux-list-back":
                curSkinProperties.AuxList.backColor = Val(rgn(1))
               Case Else:
                'INS: error-invalid object
               End Select
               
              End If
            Wend
'=======================================================================================================
        Case Else
          
          e.errNum = 3
          e.errShortDesc = "FireAMP: Script Error, Undefined region " & rgn(1)
          e.errLongDesc = "The current script file contained an invalid skin region." _
          & " Edit the script to fix this problem"
          logError e
       End Select
     Else

   End If
commentReached:
Wend
'all over, phew!
'=======================================================================================================

parseSkin = curSkinProperties
End Function

