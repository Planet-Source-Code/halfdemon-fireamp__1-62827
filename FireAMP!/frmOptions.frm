VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmOptions 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "FireAMP Preferences"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8460
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   8460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1320
      Top             =   6360
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   7080
      TabIndex        =   11
      Top             =   6360
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   6000
      TabIndex        =   10
      Top             =   6360
      Width           =   975
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Height          =   375
      Left            =   4920
      TabIndex        =   9
      Top             =   6360
      Width           =   975
   End
   Begin MSComctlLib.TreeView trvOptions 
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   8916
      _Version        =   393217
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   6
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame fraStartUP 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5175
      Left            =   2760
      TabIndex        =   2
      Top             =   960
      Width           =   5535
      Begin VB.Frame fraTypes 
         Caption         =   "File Types"
         Height          =   2415
         Left            =   360
         TabIndex        =   41
         Top             =   2640
         Visible         =   0   'False
         Width           =   4935
         Begin MSComctlLib.ListView lstTypes 
            Height          =   1935
            Left            =   120
            TabIndex        =   42
            Top             =   360
            Width           =   4695
            _ExtentX        =   8281
            _ExtentY        =   3413
            View            =   3
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            HideColumnHeaders=   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Types"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.CheckBox chkAssociate 
         Alignment       =   1  'Right Justify
         Caption         =   "Associate File types on Start"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   2280
         Width           =   4815
      End
      Begin VB.CheckBox chkDefSkin 
         Alignment       =   1  'Right Justify
         Caption         =   "Load Default Skin"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   1680
         Width           =   4815
      End
      Begin VB.CheckBox chkSplash 
         Alignment       =   1  'Right Justify
         Caption         =   "Show Splash Screen"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   1080
         Width           =   4815
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Start Up"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   4560
         TabIndex        =   19
         Top             =   300
         Width           =   765
      End
      Begin VB.Label Label9 
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   5295
      End
   End
   Begin VB.Frame fraGeneral 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5175
      Left            =   2760
      TabIndex        =   1
      Top             =   960
      Width           =   5535
      Begin VB.CheckBox chkEnableVis 
         Alignment       =   1  'Right Justify
         Caption         =   " Enable Visulaizations"
         Height          =   255
         Left            =   360
         TabIndex        =   40
         Top             =   1680
         Width           =   4815
      End
      Begin VB.CheckBox chkShowMediaTracker 
         Alignment       =   1  'Right Justify
         Caption         =   "Show MediaTracker When Minimized"
         Height          =   255
         Left            =   360
         TabIndex        =   39
         Top             =   1080
         Width           =   4815
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "General"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   4560
         TabIndex        =   21
         Top             =   300
         Width           =   765
      End
      Begin VB.Label Label15 
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   5295
      End
   End
   Begin VB.Frame fraKeyShort 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5175
      Left            =   2760
      TabIndex        =   5
      Top             =   960
      Width           =   5535
      Begin VB.ListBox lstShortCuts 
         BackColor       =   &H8000000F&
         Height          =   4110
         Left            =   120
         TabIndex        =   38
         Top             =   840
         Width           =   5295
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Keyboard Shortcuts"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   3465
         TabIndex        =   13
         Top             =   300
         Width           =   1860
      End
      Begin VB.Label Label3 
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   5295
      End
   End
   Begin VB.Frame fraAbout 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5175
      Left            =   2760
      TabIndex        =   4
      Top             =   960
      Width           =   5535
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   4095
         Left            =   120
         ScaleHeight     =   4035
         ScaleWidth      =   5235
         TabIndex        =   22
         Top             =   840
         Width           =   5295
         Begin VB.PictureBox Picture2 
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   3735
            Left            =   0
            ScaleHeight     =   3735
            ScaleWidth      =   5295
            TabIndex        =   23
            Top             =   360
            Width           =   5295
            Begin VB.Label Label21 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   240
               Left            =   2625
               TabIndex        =   27
               Top             =   1680
               Width           =   75
            End
            Begin VB.Label Label20 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Credits"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H008080FF&
               Height          =   255
               Left            =   1560
               TabIndex        =   26
               Top             =   1320
               Width           =   1695
            End
            Begin VB.Label Label19 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Version 1.0.0"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0E0FF&
               Height          =   240
               Left            =   1785
               TabIndex        =   25
               Top             =   840
               Width           =   1245
            End
            Begin VB.Label Label18 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "FireAMP!"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   26.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H005C5CCD&
               Height          =   615
               Left            =   1335
               TabIndex        =   24
               Top             =   120
               Width           =   2265
            End
         End
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "About"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   4680
         TabIndex        =   15
         Top             =   300
         Width           =   555
      End
      Begin VB.Label Label5 
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   5295
      End
   End
   Begin VB.Frame fraSkins 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5175
      Left            =   2760
      TabIndex        =   3
      Top             =   960
      Width           =   5535
      Begin VB.Frame Frame2 
         Height          =   1335
         Left            =   240
         TabIndex        =   34
         Top             =   3720
         Width           =   5055
         Begin VB.Label lblComment 
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   960
            Width           =   4815
         End
         Begin VB.Label lblCreation 
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   600
            Width           =   4815
         End
         Begin VB.Label lblAuthor 
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   240
            Width           =   4695
         End
      End
      Begin VB.ListBox lstSkins 
         Height          =   1410
         Left            =   240
         TabIndex        =   33
         Top             =   2160
         Width           =   5055
      End
      Begin VB.TextBox txtSkin 
         Height          =   375
         Left            =   240
         TabIndex        =   31
         Top             =   1200
         Width           =   3735
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "&Browse ..."
         Height          =   375
         Left            =   4080
         TabIndex        =   30
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Installed Skins:"
         Height          =   225
         Left            =   240
         TabIndex        =   32
         Top             =   1800
         Width           =   1260
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Load new Skin :"
         Height          =   225
         Left            =   240
         TabIndex        =   29
         Top             =   840
         Width           =   1305
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Skins"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   4800
         TabIndex        =   17
         Top             =   300
         Width           =   510
      End
      Begin VB.Label Label7 
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   5295
      End
   End
   Begin VB.Image Image1 
      Height          =   750
      Left            =   170
      Top             =   50
      Width           =   750
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000013&
      X1              =   0
      X2              =   8520
      Y1              =   6260
      Y2              =   6260
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000010&
      X1              =   0
      X2              =   8520
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FireAMP Preferences"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1440
      TabIndex        =   28
      Top             =   120
      Width           =   3405
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000004&
      X1              =   0
      X2              =   8520
      Y1              =   825
      Y2              =   825
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      X1              =   0
      X2              =   8520
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000006&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000000&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   8535
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'credits: my friend

'
'Created By:The Punisher
'Ideasoft, Inc.
'

Option Explicit
Private Sub chkAssociate_Click()

If chkAssociate.Value = vbChecked Then
fraTypes.Visible = True
Else
fraTypes.Visible = False
End If
End Sub

Private Sub cmdApply_Click()
Dim Options As FireAMPoptions
Options.checkAssociationsAtStartUp = chkAssociate.Value
Options.enableVisualizations = chkEnableVis.Value
Options.loadDefaultSkin = chkDefSkin.Value
Options.showMediaTracker = chkShowMediaTracker.Value
Options.showSplashScreen = chkSplash.Value

Options.MIDI = lstTypes.ListItems.Item(1).Checked
Options.WAV = lstTypes.ListItems.Item(2).Checked
Options.MP3 = lstTypes.ListItems.Item(3).Checked
Options.MPG = lstTypes.ListItems.Item(4).Checked
Options.WMA = lstTypes.ListItems.Item(5).Checked

Open App.Path & "\FireAMP.Options" For Binary Access Read Write Lock Write As #1
Put #1, , Options
Close #1
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
cmdApply_Click
Unload Me
End Sub

Private Sub Form_Load()
Dim nodOpt As Node
'Main nodes
Set nodOpt = trvOptions.Nodes.Add(, tvwNext, "Pref", "Preferences")
Set nodOpt = trvOptions.Nodes.Add(, tvwNext, "Abt", "About")
Set nodOpt = trvOptions.Nodes.Add(, tvwNext, "KbShrt", "Keyboard Shortcuts")
'Child nodes
Set nodOpt = trvOptions.Nodes.Add("Pref", tvwChild, "Gen", "General")
Set nodOpt = trvOptions.Nodes.Add("Pref", tvwChild, "StrUP", "Start Up")
Set nodOpt = trvOptions.Nodes.Add("Pref", tvwChild, "Skn", "Skins")
'Hide all frames
HideFrames
fraGeneral.Visible = True


Label18.Left = (fraAbout.Width / 2) - (Label18.Width / 2)
Label19.Left = (fraAbout.Width / 2) - (Label19.Width / 2)
Label20.Left = (fraAbout.Width / 2) - (Label20.Width / 2)
Label21.Left = (fraAbout.Width / 2) - (Label21.Width / 2)

Picture2.Height = Picture1.Height * 1.7
Picture2.Top = Picture1.Top + Picture1.Height - 1000

Label21.Caption = " Main Idea: K.Sai Krishna" _
& vbCrLf & " InterFace Designer: K.V.Rohit" _
& vbCrLf & " ~~~ " _
& vbCrLf & " ~ Version Information ~ " _
& vbCrLf & " Formats supported " _
& vbCrLf & " MIDI " _
& vbCrLf & " WAV " _
& vbCrLf & " MP1 " _
& vbCrLf & " MP2 " _
& vbCrLf & " MP3 " _
& vbCrLf & " WMA " _
& vbCrLf & " MPG " _
& vbCrLf & " DAT " _
& vbCrLf & " WMV " _
& vbCrLf & " MOV " _
& vbCrLf & " ~~~ " _
& vbCrLf & " FireScript parser version: 1.0.0 " _
& vbCrLf & " ~~~ " _
& vbCrLf & " ~ FireAMP, The Hottest Player Around ~ " _
& vbCrLf & " ~~~ " _
& vbCrLf & " ~ The End ~ "

' initialize shortcuts
lstShortCuts.AddItem "Play Clip -> Space Bar"
lstShortCuts.AddItem "Stop Clip -> S"
lstShortCuts.AddItem "Open Clip ->  O"
lstShortCuts.AddItem "Exit -> X"
lstShortCuts.AddItem "Volume up -> H"
lstShortCuts.AddItem "Volume Down -> G"
lstShortCuts.AddItem "Mute -> M"
lstShortCuts.AddItem "Minimize -> N"
lstShortCuts.AddItem "Change Visualization: Next-> B : Previous-> V"
lstShortCuts.AddItem "Full Screen/Normal video -> F"


If FSys.FileExists(App.Path & "\FireAMP.Options") Then
Open App.Path & "\FireAMP.Options" For Binary Access Read As 1
Get #1, , theOptions
Close #1
End If

 chkAssociate.Value = theOptions.checkAssociationsAtStartUp
 chkEnableVis.Value = theOptions.enableVisualizations
 chkDefSkin.Value = theOptions.loadDefaultSkin
 chkShowMediaTracker.Value = theOptions.showMediaTracker
 chkSplash.Value = theOptions.showSplashScreen
 

lstTypes.ListItems.Add , , "MIDI Sequences (MID,RMI)"
lstTypes.ListItems.Add , , "WAVE Audio (WAV)"
lstTypes.ListItems.Add , , "MPEG Audio (MP3, MP2, MP1)"
lstTypes.ListItems.Add , , "MPEG Video (MPG, MPEG, MPE)"
lstTypes.ListItems.Add , , "Windows Media (WMA, WMV)"

lstTypes.ColumnHeaders(1).Width = lstTypes.Width

If theOptions.MIDI Then lstTypes.ListItems.Item(1).Checked = True
If theOptions.MP3 Then lstTypes.ListItems.Item(3).Checked = True
If theOptions.MPG Then lstTypes.ListItems.Item(4).Checked = True
If theOptions.WAV Then lstTypes.ListItems.Item(2).Checked = True
If theOptions.WMA Then lstTypes.ListItems.Item(5).Checked = True


End Sub

Private Sub Timer1_Timer()
Picture2.Top = Picture2.Top - 50
If Picture2.Top < -Picture2.Height Then Picture2.Top = Picture1.Top + Picture1.Height
End Sub

Private Sub trvOptions_NodeClick(ByVal Node As MSComctlLib.Node)
'Note you can call "HideFrames" in each case also
HideFrames
Timer1.Enabled = False

Select Case Node.Key
Case "Gen", "Pref"
    fraGeneral.Visible = True
Case "StrUP"
    fraStartUP.Visible = True
    fraTypes.Visible = chkAssociate.Value
Case "Skn"
    fraSkins.Visible = True
Case "Abt"
    fraAbout.Visible = True
       
    Timer1.Enabled = True
Case "KbShrt"
    fraKeyShort.Visible = True
Case Default

End Select
End Sub

Private Sub HideFrames()
Dim Control As Object
    For Each Control In Me
        If TypeOf Control Is Frame Then
            Control.Visible = False
        End If
    Next
End Sub
