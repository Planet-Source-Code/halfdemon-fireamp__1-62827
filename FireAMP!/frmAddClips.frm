VERSION 5.00
Begin VB.Form frmAddClips 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Add Media Clips"
   ClientHeight    =   8940
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10500
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8940
   ScaleWidth      =   10500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   8400
      Top             =   4560
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000080&
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   120
      ScaleHeight     =   195
      ScaleWidth      =   5820
      TabIndex        =   6
      Top             =   8400
      Visible         =   0   'False
      Width           =   5850
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   0
         Picture         =   "frmAddClips.frx":0000
         ScaleHeight     =   195
         ScaleWidth      =   15
         TabIndex        =   24
         Top             =   0
         Width           =   15
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8760
      TabIndex        =   5
      Top             =   8400
      Width           =   1455
   End
   Begin VB.CommandButton cmdAddPL 
      Caption         =   "&Add Clips"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6960
      TabIndex        =   4
      Top             =   8400
      Width           =   1575
   End
   Begin VB.Frame fraAdd 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6735
      Left            =   0
      TabIndex        =   7
      Top             =   1440
      Width           =   10455
      Begin VB.Frame Frame5 
         Caption         =   "Built up Playlist:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6375
         Left            =   4920
         TabIndex        =   15
         Top             =   240
         Width           =   5415
         Begin VB.ListBox lstPL 
            Height          =   3630
            Left            =   120
            TabIndex        =   23
            Top             =   360
            Width           =   5175
         End
         Begin VB.ListBox lstPaths 
            Height          =   2580
            Left            =   480
            TabIndex        =   22
            Top             =   840
            Visible         =   0   'False
            Width           =   2895
         End
         Begin VB.Frame Frame6 
            Caption         =   "Options:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2175
            Left            =   120
            TabIndex        =   16
            Top             =   4080
            Width           =   5175
            Begin VB.CommandButton Command5 
               Caption         =   "&A&dd All Files from Dir."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               TabIndex        =   43
               Top             =   1320
               Width           =   1935
            End
            Begin VB.CheckBox chkCheck 
               Caption         =   "Donot Scan Files (makes the addition faster)"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   240
               TabIndex        =   21
               Top             =   1800
               Width           =   3975
            End
            Begin VB.CommandButton Command1 
               Caption         =   "&Remove Selected"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               TabIndex        =   20
               Top             =   360
               Width           =   1935
            End
            Begin VB.CommandButton Command2 
               Caption         =   "Remove &All"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               TabIndex        =   19
               Top             =   840
               Width           =   1935
            End
            Begin VB.CommandButton Command3 
               Caption         =   "ScanDisk for Media"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   3000
               TabIndex        =   18
               Top             =   360
               Width           =   1935
            End
            Begin VB.CommandButton Command4 
               Caption         =   "Edit MP3 ID3 Tag"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   3000
               TabIndex        =   17
               Top             =   840
               Width           =   1935
            End
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "File Browser:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6375
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   4335
         Begin VB.Frame Frame2 
            Caption         =   "Drives:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   4095
            Begin VB.DriveListBox Drive1 
               Height          =   330
               Left            =   240
               TabIndex        =   14
               Top             =   240
               Width           =   3615
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Directories:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2295
            Left            =   120
            TabIndex        =   11
            Top             =   1080
            Width           =   4095
            Begin VB.DirListBox Dir1 
               Height          =   1770
               Left            =   120
               TabIndex        =   12
               Top             =   240
               Width           =   3855
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Files:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2775
            Left            =   120
            TabIndex        =   9
            Top             =   3480
            Width           =   4095
            Begin VB.FileListBox File1 
               Height          =   2400
               Left            =   120
               Pattern         =   "*.mid;*.rmi;*.mp3;*.mp2;*.mp1;*.wav;*.wma;*.wmv;*.mpg;*.mpeg;*.dat"
               TabIndex        =   10
               Top             =   240
               Width           =   3855
            End
         End
      End
   End
   Begin VB.Frame fraScan 
      Caption         =   "Scan Disk for Media"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6735
      Left            =   0
      TabIndex        =   26
      Top             =   1440
      Width           =   10455
      Begin VB.Frame Frame9 
         Caption         =   "Select Drive to Scan"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3240
         TabIndex        =   41
         Top             =   840
         Width           =   4095
         Begin VB.DriveListBox Drive2 
            Height          =   330
            Left            =   240
            TabIndex        =   42
            Top             =   240
            Width           =   3735
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Scan For"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   3240
         TabIndex        =   35
         Top             =   1680
         Width           =   4095
         Begin VB.CheckBox chkMIDI 
            Caption         =   "MIDI Sequences (MID,RMI)"
            Height          =   255
            Left            =   240
            TabIndex        =   40
            Top             =   360
            Width           =   2175
         End
         Begin VB.CheckBox chkWAV 
            Caption         =   " WAVE Audio"
            Height          =   255
            Left            =   240
            TabIndex        =   39
            Top             =   720
            Width           =   2175
         End
         Begin VB.CheckBox chkMP3 
            Caption         =   "MPEG Audio (MP1, MP2, MP3)"
            Height          =   255
            Left            =   240
            TabIndex        =   38
            Top             =   1080
            Width           =   2535
         End
         Begin VB.CheckBox chkMPEG 
            Caption         =   "MPEG Videos (MPEG, MPG, MPE)"
            Height          =   255
            Left            =   240
            TabIndex        =   37
            Top             =   1440
            Width           =   3015
         End
         Begin VB.CheckBox chkWMA 
            Caption         =   " Windows Media (WMA, WMV)"
            Height          =   255
            Left            =   240
            TabIndex        =   36
            Top             =   1800
            Width           =   2655
         End
      End
      Begin VB.Frame fraProgress 
         Caption         =   "Scan Progress"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   240
         TabIndex        =   30
         Top             =   5160
         Visible         =   0   'False
         Width           =   10095
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00000080&
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   960
            ScaleHeight     =   195
            ScaleWidth      =   7500
            TabIndex        =   31
            Top             =   480
            Width           =   7530
            Begin VB.PictureBox lblBar 
               BorderStyle     =   0  'None
               Height          =   195
               Left            =   -15
               Picture         =   "frmAddClips.frx":4C6E
               ScaleHeight     =   195
               ScaleWidth      =   15
               TabIndex        =   32
               Top             =   0
               Width           =   15
            End
         End
         Begin VB.Label lblPath 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   34
            Top             =   840
            Width           =   7935
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblCount 
            Caption         =   "Found: 0 Clips"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   8280
            TabIndex        =   33
            Top             =   840
            Width           =   1695
         End
      End
      Begin VB.CommandButton cmdScan 
         Caption         =   "&Scan Away!"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   29
         Top             =   4440
         Width           =   1335
      End
      Begin VB.CommandButton cmdDone 
         Caption         =   "&Done"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4680
         TabIndex        =   28
         Top             =   4440
         Width           =   1215
      End
      Begin VB.CommandButton cmdBack 
         Caption         =   "< Back"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6000
         TabIndex        =   27
         Top             =   4440
         Width           =   1335
      End
   End
   Begin VB.Label lblProgress 
      Caption         =   "Progress... "
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   8640
      Visible         =   0   'False
      Width           =   5895
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000013&
      X1              =   0
      X2              =   10560
      Y1              =   8300
      Y2              =   8300
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000010&
      X1              =   0
      X2              =   10560
      Y1              =   8280
      Y2              =   8280
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "[Double Click on a file name to include it in list]"
      Height          =   210
      Left            =   6720
      TabIndex        =   3
      Top             =   960
      Width           =   3270
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000013&
      X1              =   0
      X2              =   10560
      Y1              =   1335
      Y2              =   1335
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   0
      X2              =   10560
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FireAMP!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8640
      TabIndex        =   2
      Top             =   480
      Width           =   1365
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Add Media Clips"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   3765
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1.00100e5
   End
End
Attribute VB_Name = "frmAddClips"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim sSize As Double, tSize As Double, nCount As Long
Dim br As Boolean
Dim running As Boolean

Private Sub cmdAddPL_Click()

fraAdd.Enabled = False
If lstPL.ListCount > 30 And chkCheck.Value = 0 Then
Dim a As VbMsgBoxResult
a = MsgBox("It appears that you have a lot of items in your playlist. (" & lstPL.ListCount & " clips in fact) " & vbNewLine & "Are you sure you want to scan them before adding? [Scanning takes a long time!]", vbInformation + vbYesNo, "Scan Files?")
If a = vbYes Then
chkCheck.Value = 0
Else
chkCheck.Value = 1
End If
End If
chkCheck.Refresh

On Error GoTo errHandle

Dim i As Integer, addCount As Integer


Picture1.Visible = True
lblProgress.Visible = True
frmFirePL.lstPaths.Clear
frmFirePL.lstPL.ListItems.Clear

For i = 0 To lstPL.ListCount - 1
DoEvents
If Trim(lstPL.List(i)) = "" Then GoTo JMP
If PlayClip(lstPaths.List(i), True) Then
frmFirePL.lstPL.ListItems.Add , , lstPL.List(i)
frmFirePL.lstPaths.AddItem lstPaths.List(i)
addCount = addCount + 1

End If
Picture2.Width = Picture1.ScaleWidth * ((i + 1) / lstPL.ListCount)
lblProgress.Caption = "Progress ... [ Scanned " & i + 1 & " clip(s) of " & lstPL.ListCount & ", Added: " & addCount & " ]"
JMP:
Next

MsgBox "Addition Complete!" & vbNewLine & "Total files: " & lstPL.ListCount & vbNewLine & "Added files: " & addCount & vbNewLine & "Rejected files: " & lstPL.ListCount - addCount, vbInformation + vbOKOnly, "Statistics"
errHandle:
Unload Me
End Sub

Private Sub cmdBack_Click()
fraScan.Visible = False
fraAdd.Visible = True
End Sub

Private Sub cmdCancel_Click()
If running Then
br = True
End If
Unload Me
Set frmAddClips = Nothing
End Sub


Private Sub cmdDone_Click()
fraScan.Visible = False
fraAdd.Visible = True

cmdScan.Enabled = True
cmdBack.Enabled = True
cmdAddPL.Enabled = True
fraProgress.Visible = False
End Sub

Private Sub cmdScan_Click()

Dim chk As Control, Count As Integer
For Each chk In Me
If TypeOf chk Is CheckBox Then
If chk.Value = 1 And chk <> chkCheck Then Count = Count + 1
End If

Next

If Count = 0 Then
Dim e As ErrStruct
e.errNum = 4
e.errShortDesc = "No media type selected!"
e.errLongDesc = "To scan a drive for media, select a media type"
logError e
Exit Sub
End If
running = True
Dim Path As String
Path = Drive2.Drive & "\"
nCount = 0
cmdScan.Enabled = False
cmdBack.Enabled = False
DoEvents
fraProgress.Visible = True
sSize = 0
tSize = FSys.GetDrive(Path).AvailableSpace
scanFolder Path
cmdDone.Enabled = True
Beep
running = False

End Sub

Private Sub Command1_Click()
On Error Resume Next
lstPL.RemoveItem lstPL.ListIndex
lstPaths.RemoveItem lstPL.ListIndex

End Sub

Private Sub Command2_Click()
On Error Resume Next
lstPL.Clear
lstPaths.Clear
End Sub

Private Sub Command3_Click()
fraAdd.Visible = False
fraScan.Visible = True
cmdAddPL.Enabled = False
End Sub

Private Sub Command5_Click()

Dim i As Integer
For i = 0 To File1.ListCount - 1
File1.ListIndex = i
File1_DblClick
Next i
End Sub

Private Sub Command4_Click()
frmTagEditor.FILENAME = lstPaths.List(lstPL.ListIndex)
frmTagEditor.Show vbModal
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive & "\"
End Sub

Private Sub File1_DblClick()
Dim theFile As String
If Right(File1.Path, 1) = "\" Then
 theFile = File1.Path & File1.FILENAME
Else
 theFile = File1.Path & "\" & File1.FILENAME
End If
lstPL.AddItem getFileName(theFile)
lstPaths.AddItem theFile
           If lstPL.ListCount <> 0 Then
           cmdAddPL.Enabled = True
           Else
           cmdAddPL.Enabled = False
           End If

End Sub


Sub scanFolder(FolderSpec As String)
On Error GoTo e
DoEvents
Dim i As Integer

Dim thisFolder As Folder
Dim sFolders As Folders
Dim fileItem As File, folderItem As Folder
Dim allFiles As Files

Set thisFolder = FSys.GetFolder(FolderSpec)
Set sFolders = thisFolder.SubFolders
Set allFiles = thisFolder.Files


For Each folderItem In sFolders
DoEvents
lblPath.Caption = "Scanning -- " & folderItem.Path
                    
If br Then Exit Sub
scanFolder (folderItem.Path)

Next

For Each fileItem In allFiles

If isMediaFile(fileItem.Path) Then
nCount = nCount + 1
lstPaths.AddItem fileItem.Path
lstPL.AddItem getFileName(fileItem.Path)
If br Then Exit Sub
End If
sSize = sSize + fileItem.Size
Next
DoEvents
lblBar.Width = Picture3.ScaleWidth * (sSize / tSize)
lblCount.Caption = "Found: " & nCount & " Clips"
Exit Sub
e:
If Err.Number = 76 Then MsgBox "The specified file, folder or drive was not accessible" & vbCrLf & "Please try again", vbOKOnly + vbInformation, "Path not found"

End Sub

Function isMediaFile(FilePath As String) As Boolean
Dim isValid As Boolean
Dim ext As String


ext = getExtension(FilePath)

If (ext = "mid" Or ext = "rmi") And CBool(chkMIDI.Value) Then
isValid = True
ElseIf (ext = "wav") And CBool(chkWAV.Value) Then
isValid = True
ElseIf (ext = "mp3" Or ext = "mp1" Or ext = "mp1") And CBool(chkMP3.Value) Then
isValid = True
ElseIf (ext = "mpeg" Or ext = "mpg" Or ext = "dat") And CBool(chkMPEG.Value) Then
isValid = True
ElseIf (ext = "wma" Or ext = "wmv") And CBool(chkWMA.Value) Then
isValid = True
Else
isValid = False
End If
isMediaFile = isValid
End Function


Private Sub Form_Load()

If frmFirePL.lstPaths.ListCount > 0 Then
Dim i As Integer
For i = 0 To frmFirePL.lstPaths.ListCount - 1
 lstPL.AddItem getFileName(frmFirePL.lstPaths.List(i))
 lstPaths.AddItem frmFirePL.lstPaths.List(i)
Next
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
Set frmAddClips = Nothing
End Sub

Private Sub lstPL_Click()
If getExtension(lstPaths.List(lstPL.ListIndex)) = "mp3" Then
Command4.Enabled = True
Else
Command4.Enabled = False
End If
End Sub

Private Sub Timer1_Timer()
If lstPL.ListCount > 0 Then
cmdAddPL.Enabled = True
Else
cmdAddPL.Enabled = False
End If

End Sub
