VERSION 5.00
Begin VB.Form frmTagEditor 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "FireAMP Tag Editor"
   ClientHeight    =   7665
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9030
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   9030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      Height          =   735
      Left            =   120
      TabIndex        =   38
      Top             =   960
      Width           =   8775
      Begin VB.TextBox txtFilename 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   270
         Width           =   6255
      End
      Begin VB.CommandButton cmdParse 
         Caption         =   "Parse to Tag"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7320
         TabIndex        =   39
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Filename"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   41
         Top             =   315
         Width           =   780
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Other Options"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   4560
      TabIndex        =   37
      Top             =   6000
      Width           =   2895
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "&Quit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   36
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton cmdCommit 
      Caption         =   "&Commit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   35
      Top             =   6600
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      Caption         =   "ID3 v2.0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   4560
      TabIndex        =   20
      Top             =   1920
      Width           =   4365
      Begin VB.CommandButton cmdCopyID3v1 
         Caption         =   "Copy from ID3 v1 Tag"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2040
         TabIndex        =   43
         Top             =   3390
         Width           =   1935
      End
      Begin VB.TextBox txtSongTitle2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         TabIndex        =   27
         Top             =   480
         Width           =   2775
      End
      Begin VB.TextBox txtArtist2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         TabIndex        =   26
         Top             =   960
         Width           =   2775
      End
      Begin VB.TextBox txtAlbum2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         TabIndex        =   25
         Top             =   1440
         Width           =   2775
      End
      Begin VB.TextBox txtComments2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         TabIndex        =   24
         Top             =   1920
         Width           =   2775
      End
      Begin VB.TextBox txtYear2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         TabIndex        =   23
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox txtTrack2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3360
         TabIndex        =   22
         Top             =   2400
         Width           =   735
      End
      Begin VB.ComboBox cboGenre2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1320
         TabIndex        =   21
         Top             =   2880
         Width           =   2775
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Song Title"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   34
         Top             =   510
         Width           =   825
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Artist"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   33
         Top             =   1005
         Width           =   405
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Album"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   32
         Top             =   1485
         Width           =   525
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Comments"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   225
         TabIndex        =   31
         Top             =   1965
         Width           =   930
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Year"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   30
         Top             =   2445
         Width           =   375
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Genre"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   29
         Top             =   2925
         Width           =   510
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Track #"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2640
         TabIndex        =   28
         Top             =   2445
         Width           =   600
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "File Properties"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   15
      Top             =   6000
      Width           =   4335
      Begin VB.CheckBox chkSystem 
         Caption         =   "System File"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   19
         Top             =   960
         Width           =   1335
      End
      Begin VB.CheckBox chkArchive 
         Caption         =   "Archive"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   18
         Top             =   480
         Width           =   975
      End
      Begin VB.CheckBox chkHidden 
         Caption         =   "Hidden"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   960
         Width           =   975
      End
      Begin VB.CheckBox chkReadOnly 
         Caption         =   "Read Only"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "ID3 v1.0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   4365
      Begin VB.CommandButton cmdCopyID3v2 
         Caption         =   "Copy from ID3 v2 Tag"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2025
         TabIndex        =   42
         Top             =   3375
         Width           =   1935
      End
      Begin VB.ComboBox cboGenre 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "frmTagEditor.frx":0000
         Left            =   1320
         List            =   "frmTagEditor.frx":0002
         TabIndex        =   14
         Text            =   "cboGenre"
         Top             =   2880
         Width           =   2775
      End
      Begin VB.TextBox txtNumber 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3360
         TabIndex        =   12
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox txtYear 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         TabIndex        =   11
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox txtComments 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         TabIndex        =   10
         Top             =   1920
         Width           =   2775
      End
      Begin VB.TextBox txtAlbum 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         TabIndex        =   9
         Top             =   1440
         Width           =   2775
      End
      Begin VB.TextBox txtArtist 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         TabIndex        =   8
         Top             =   960
         Width           =   2775
      End
      Begin VB.TextBox txtSongTitle 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         TabIndex        =   7
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Track #"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2640
         TabIndex        =   13
         Top             =   2445
         Width           =   600
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Genre"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   6
         Top             =   2925
         Width           =   510
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Year"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   5
         Top             =   2445
         Width           =   375
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Comments"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   225
         TabIndex        =   4
         Top             =   1965
         Width           =   930
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Album"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   3
         Top             =   1485
         Width           =   525
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Artist"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   2
         Top             =   1005
         Width           =   405
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Song Title"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   1
         Top             =   510
         Width           =   825
      End
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000004&
      X1              =   0
      X2              =   9600
      Y1              =   825
      Y2              =   825
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      X1              =   0
      X2              =   9600
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   9615
   End
End
Attribute VB_Name = "frmTagEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'thanks to Punisher, IdeaSoft Inc,
' ID3v2 need implementation

Dim currentTag As tagMP3ID3V1
Dim NUM As Byte
Public FILENAME As String

Private Sub cmdCommit_Click()
    
    SaveTagV1 (Trim(txtFilename.Text))
    
    Unload Me
    
End Sub

Private Sub cmdCopyID3v1_Click()

    txtSongTitle2.Text = Trim(txtSongTitle.Text)
    txtArtist2.Text = Trim(txtArtist.Text)
    txtAlbum2.Text = Trim(txtAlbum.Text)
    txtComments2.Text = Trim(txtComments.Text)
    txtYear2.Text = Trim(txtYear.Text)
    txtTrack2.Text = Trim(txtNumber.Text)
    cboGenre2.Text = cboGenre.Text
    
End Sub

Private Sub cmdCopyID3v2_Click()

    txtSongTitle.Text = Trim(txtSongTitle2.Text)
    txtArtist.Text = Trim(txtArtist2.Text)
    txtAlbum.Text = Trim(txtAlbum2.Text)
    txtComments.Text = Trim(txtComments2.Text)
    txtYear.Text = Trim(txtYear2.Text)
    txtNumber.Text = Trim(txtTrack2.Text)
    cboGenre.Text = cboGenre2.Text
    
End Sub

Private Sub cmdParse_Click()

On Error Resume Next

Dim FLName As String
Dim N As Long
    FLName = txtFilename.Text
    
    N = InStrRev(FLName, "\") - 1
    N = (Len(txtFilename.Text) - N) - 1
    FLName = Right(FLName, N)
    N = Len(FLName) - 4
    FLName = Left(FLName, N)
        
    N = InStrRev(FLName, "-") - 1
    
    txtArtist.Text = Trim(Left(FLName, N))
    txtArtist2.Text = txtArtist.Text
    
    N = Len(FLName) - Len(txtArtist.Text)
    FLName = Trim(Right(FLName, N))
    N = InStrRev(FLName, "-")
    N = Len(FLName) - N
    
    txtSongTitle.Text = Trim(Right(FLName, N))
    txtSongTitle2.Text = txtSongTitle.Text
    
End Sub

Private Sub cmdQuit_Click()

    Unload Me
    
End Sub

Private Sub Form_Load()

    txtFilename.Text = FILENAME
    'Get file attributes
    
    Call GetAttributes(Trim(txtFilename.Text))
    
    For NUM = 0 To 147
        cboGenre.AddItem getGenre(NUM)
        cboGenre2.AddItem getGenre(NUM)
    Next NUM
    'Read file tag
    Call ReadTagV1(Trim(txtFilename.Text))
           
End Sub

Private Function ReadTagV1(FILENAME As String)

    currentTag = readMP3Tag(FILENAME)
    
    txtSongTitle.Text = currentTag.Title
    txtArtist.Text = currentTag.Artist
    txtAlbum.Text = currentTag.Album
    txtComments.Text = currentTag.Comment
    txtYear.Text = currentTag.Year
    txtNumber.Text = currentTag.Track
    cboGenre.Text = getGenre(currentTag.Genre)
    
End Function

Private Function SaveTagV1(FILENAME As String)
On Error Resume Next
    currentTag.Title = Trim(txtSongTitle.Text)
    currentTag.Artist = Trim(txtArtist.Text)
    currentTag.Album = Trim(txtAlbum.Text)
    currentTag.Comment = Trim(txtComments.Text)
    currentTag.Genre = cboGenre.ListIndex
    currentTag.Year = Trim(txtYear.Text)
    currentTag.Track = Trim(txtNumber.Text)
    
    Call writeTag(currentTag, FILENAME)

End Function

Private Function GetAttributes(FILENAME As String)

    Dim ATTR As Integer
    ATTR = GetAttr(FILENAME)
    chkReadOnly = IIf((ATTR And vbReadOnly) = vbReadOnly, 1, 0)
    chkHidden = IIf((ATTR And vbHidden) = vbHidden, 1, 0)
    chkSystem = IIf((ATTR And vbSystem) = vbSystem, 1, 0)
    chkArchive = IIf((ATTR And vbArchive) = vbArchive, 1, 0)

End Function
