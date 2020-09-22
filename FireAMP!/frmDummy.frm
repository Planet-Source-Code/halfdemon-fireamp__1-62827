VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmDummy 
   Caption         =   "Form1"
   ClientHeight    =   8400
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9870
   LinkTopic       =   "Form1"
   ScaleHeight     =   560
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   658
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cd1 
      Left            =   5760
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFireAMP 
      Caption         =   "FireAMP"
      WindowList      =   -1  'True
      Begin VB.Menu mnuAbout 
         Caption         =   "About FireAMP!"
      End
      Begin VB.Menu mnuNull0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowPlaylist 
         Caption         =   "Show Playlist (S)"
      End
      Begin VB.Menu mnuOpenMedia 
         Caption         =   "Open Media (O)"
      End
      Begin VB.Menu mnuPreferences 
         Caption         =   "Preferences (D)"
      End
      Begin VB.Menu mnuFullScreen 
         Caption         =   "Full Screen Video (F)"
      End
      Begin VB.Menu mnuChangeSkin 
         Caption         =   "Change Skin (C)"
      End
      Begin VB.Menu mnuNull1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit (X)"
      End
   End
   Begin VB.Menu mnuPlaylist 
      Caption         =   "Playlist"
      Begin VB.Menu mnuTagEdit 
         Caption         =   "&Edit Tag (E)"
      End
      Begin VB.Menu mnuAddClips 
         Caption         =   "Add Clips..."
      End
      Begin VB.Menu mnuSearch 
         Caption         =   "Find Clip in Playlist ..."
      End
      Begin VB.Menu mnuNull2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpenPL 
         Caption         =   "Open Playlist"
      End
      Begin VB.Menu mnuSavePL 
         Caption         =   "Save Playlist"
      End
   End
End
Attribute VB_Name = "frmDummy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuAbout_Click()

    Abt = True
    frmOptions.Show vbModal
    frmOptions.Timer1.Enabled = True

End Sub

Private Sub mnuAddClips_Click()
frmAddClips.Show vbModal
End Sub

Private Sub mnuChangeSkin_Click()
cd1.FILENAME = ""
cd1.Filter = "FireAMP! Skins (*.cfs)|*.cfs"
cd1.ShowOpen
If cd1.FILENAME <> "" Then
renderSkin cd1.FILENAME
End If

End Sub

Private Sub mnuExit_Click()

End

End Sub

Private Sub mnuOpenPL_Click()
cd1.FILENAME = ""
cd1.Filter = "FireAMP! Playlists(*.fpl)|*.fpl"
cd1.ShowOpen
If cd1.FILENAME <> "" Then
openPlayList frmFirePL.lstPL, frmFirePL.lstPaths, cd1.FILENAME
frmFirePL.lblCaption.Caption = "Playlist- " & getFileName(cd1.FILENAME)
End If
End Sub

Private Sub mnuPreferences_Click()

    frmOptions.Show vbModal

End Sub

Private Sub mnuSavePL_Click()
cd1.FILENAME = ""
cd1.Filter = "FireAMP! Playlists(*.fpl)|*.fpl"
cd1.ShowSave
If cd1.FILENAME <> "" Then
savePlayList frmFirePL.lstPaths, cd1.FILENAME
End If

End Sub

Private Sub mnuSearch_Click()
frmSearch.Show
End Sub

Private Sub mnuTagEdit_Click()
If frmFirePL.lstPL.ListItems.Count > 0 Then
 frmTagEditor.FILENAME = frmFirePL.lstPaths.List(frmFirePL.lstPL.SelectedItem.Index - 1)
    frmTagEditor.Visible = True
End If
    
End Sub
