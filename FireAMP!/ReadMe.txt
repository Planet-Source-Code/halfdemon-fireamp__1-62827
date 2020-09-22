FireAMP! - The Media Player, Rewinded

Author: K.Sai Krishna <k.saikrishna@rediffmail.com>

FireAMP! readme
---------------

FireAMP! is a (to be) versatile media player written in VB6 Enterprise (compatible with VB6 Pro and VB5). As usual, this player is a FreeWare.

Features (Beta 1)
-----------------

* FirePlaylists: FPL
* Visualizations: 14 of them, extensible
* Video Playback
* Full Screen Video
* Skins
* MP3 tag reader/writer
* Keyboard shortcuts
* Title of song scrolls if it is too long. The scrolling is not continuous, kind of like   RealPlayer's

Features to be added: 
* FVS: FireAMP! Visualization Studio
* Advanced, flexible skinning
* MIDI data reading, WMA tag reading
* Media Library

In short, by the time FireAMP! is fully built,

  FireAMP = WinAMP + Windows Media Player 9 + Real Player 10

should be satisfied!

General Notes:
--------------

[1] Key Board Shortcuts:

Space Bar: Play
S: Stop
O: Open
X: Exit
H: Volume up
G: Volume down
M: Mute
N: Minimize
F: Full Screen Video
V/B: Change vis. (Forward/Backward)
C: Play Video CD

[2] Visualizations:

FireAMP!, as of now uses 'Hard-coded' vis. I'am working on an extensible sytem like AVS using the Microsoft Script Control. The flip-side is that it is quite slow. Anyway, even if it is slow, it uses only half the CPU power of Winamp AVS.

Feel free to mess around with the existing vis. and if you feel that you have created a new, cool vis., send me the code and I'll see if I can include it in the Hard-coded vis. list!

[3] Skinning:

FireAMP! Skinning is quite crude as of now. I've included a skin called 'Platinum' in a file called 'Platinum.cfs'. Load the skin and see the difference.

Creating skins is as of now a quite a long process. After you have loaded a skin in FireAMP!, open the 'Temp' directory in the App path and you will find a number of files; these are the skin elements. You can also see a file with extension '.fss'. This is the Skin Specification File. You can open it using NotePad. I can't porvide with a complete documentation of FireScript (the code in the fss file) as of now; but maybe in future! The fss file is quite self explainatory.

Use the Skin Compiler utility (Source code included) to compile the skin. Make sure all the files: fss file and images are in the same directory. The compiler makes a file with extension '.cfs' which is the Skin to be used.

[4] Finally,

I hate using third party dll's like FMOD or Impulse Studio. I prefer to write my own so that I get all the power, performance and features I'll ever need. If you are planning to extend FireAMP!, no greasy dll's please...

I also take care of CPU usage too. Any further advancements must not hog the CPU.
The FFT code may need cleaning up. I donot know about FFT as of know, so I just 'cut and pasted' Murphy McCauley (MurphyMc@Concentric.NET)'s VB FFT code based on D.Cross's FFT code. Check out www.fullspectrum.com for more details.

I'll learn all about FFT next semester in college, so I think I can re-write the code.

[5] Credits

Murphy McCauley (MurphyMc@Concentric.NET): for his VB FFT code
K.V.Rohit (KV.Rohit@rediffmail.com): For the Tag reader/writer and for the 'Spectroscope' vis. He also designed the interface for Tag Editor and Options Page. He is currently my WebSite builder and Skin Artist.

Be sure to visit www.voidmain.cjb.net for more great VB code!

That's All Folks!