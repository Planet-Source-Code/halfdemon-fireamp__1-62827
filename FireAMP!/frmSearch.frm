VERSION 5.00
Begin VB.Form frmSearch 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Search Playlist"
   ClientHeight    =   1065
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1065
   ScaleWidth      =   3510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&Restart"
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
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   975
   End
   Begin VB.ListBox List1 
      Height          =   450
      Left            =   480
      TabIndex        =   2
      Top             =   2640
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "0 Items found"
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
      Left            =   1920
      TabIndex        =   4
      Top             =   600
      Width           =   1455
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim LST As ListItem, searched As Boolean, j As Integer

Private Sub Command1_Click()
If Not searched Then
Dim i As Integer
For i = 1 To frmFirePL.lstPL.ListItems.Count
If InStr(1, frmFirePL.lstPL.ListItems.Item(i).Text, Text1.Text, vbTextCompare) Then
frmFirePL.lstPL.ListItems.Item(i).foreColor = vbWhite
List1.AddItem i
End If
Next
Label1.Caption = List1.ListCount & " Items found"
Else
frmFirePL.lstPL.ListItems.Item(Val(List1.List(j))).EnsureVisible
frmFirePL.lstPL.ListItems.Item(Val(List1.List(j))).foreColor = vbGreen
If j > 0 Then frmFirePL.lstPL.ListItems.Item(Val(List1.List(j - 1))).foreColor = vbWhite
j = j + 1
If j > List1.ListCount - 1 Then j = 0
End If
searched = True

End Sub

Private Sub Command2_Click()
Unload Me
frmSearch.Show
End Sub

Private Sub Form_Load()
searched = False
j = 0
Dim Count As Integer
For Count = 1 To frmFirePL.lstPL.ListItems.Count
frmFirePL.lstPL.ListItems.Item(Count).foreColor = plColor
Next Count

End Sub

