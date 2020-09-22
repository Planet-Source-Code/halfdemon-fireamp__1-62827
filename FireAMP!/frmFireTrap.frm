VERSION 5.00
Begin VB.Form frmFireTrap 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "FireAMP ErrorTrap"
   ClientHeight    =   1515
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4395
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1515
   ScaleWidth      =   4395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Help"
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   1080
      Width           =   855
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000013&
      BorderStyle     =   6  'Inside Solid
      X1              =   0
      X2              =   4440
      Y1              =   980
      Y2              =   980
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   0
      X2              =   4440
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label lblNum 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   720
      TabIndex        =   4
      Top             =   120
      Width           =   1980
   End
   Begin VB.Label lblReason 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   4215
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmFireTrap.frx":0000
      Top             =   240
      Width           =   480
   End
   Begin VB.Label lblError 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   3615
   End
End
Attribute VB_Name = "frmFireTrap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim State As Boolean
Dim t As Long
Private Declare Function MessageBeep Lib "user32" (ByVal wType As Long) As Long

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
State = Not State
If State Then
t = lblReason.Top
Height = Height + lblReason.Height + 100
Command1.Top = Command1.Top + lblReason.Height + 100
Command2.Top = Command1.Top
lblReason.Top = Line2.Y1 + 20
Else
lblReason.Top = t
Height = Height - lblReason.Height - 100
Command1.Top = Command1.Top - lblReason.Height - 100
Command2.Top = Command1.Top
End If
End Sub

Private Sub Form_Load()
MessageBeep vbError ' beep away...
State = False
End Sub
