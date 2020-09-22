VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMediaLib 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FireAMP Media Library"
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   9975
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView lstMediaLib 
      Height          =   5295
      Left            =   2640
      TabIndex        =   1
      Top             =   1080
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   9340
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
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
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "SngTitle"
         Text            =   "Songtitle"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Artist"
         Text            =   "Artist"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "Album"
         Text            =   "Album"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "Year"
         Text            =   "Year"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "Comment"
         Text            =   "Comment"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Key             =   "Track"
         Text            =   "Track #"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Key             =   "Filename"
         Text            =   "Filename"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.TreeView trvMediaLib 
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   9340
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
   Begin VB.Line Line2 
      BorderColor     =   &H80000000&
      X1              =   0
      X2              =   10200
      Y1              =   820
      Y2              =   820
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      X1              =   0
      X2              =   10200
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000000&
      BorderStyle     =   0  'Transparent
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   10215
   End
End
Attribute VB_Name = "frmMediaLib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim trvMediaNode As Node
Set trvMediaNode = trvMediaLib.Nodes.Add(, tvwNext, "Lib", "Library")
Set trvMediaNode = trvMediaLib.Nodes.Add(, tvwNext, "Play", "Playlist(s)")
Set trvMediaNode = trvMediaLib.Nodes.Add("Play", tvwChild, "usrPlay1", "Playlist1")
Set trvMediaNode = trvMediaLib.Nodes.Add("Play", tvwChild, "usrPlay2", "Playlist2")
Set trvMediaNode = trvMediaLib.Nodes.Add("Play", tvwChild, "usrPlay3", "Playlist3")
End Sub
