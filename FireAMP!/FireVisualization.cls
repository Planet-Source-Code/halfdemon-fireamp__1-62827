VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FireVisualization"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public thePictureBox As PictureBox
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Public Sub Line(X1 As Integer, Y1 As Integer, X2 As Integer, Y2 As Integer, Color As Long)
If isValid(thePictureBox) Then
'thePictureBox.Line (X1, Y1)-(X2, Y2), Color
Rectangle thePictureBox.hdc, X1, Y1, 1 + X2, Y2
Else
' error
End If
End Sub

Public Sub Point(X As Integer, Y As Integer, Color As Long)
If isValid(thePictureBox) Then
thePictureBox.PSet (X, Y), Color
Else
' error
End If
End Sub


Private Function isValid(obj As Object)
isValid = (obj = Empty)
End Function
