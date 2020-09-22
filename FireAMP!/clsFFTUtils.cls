VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FFTutils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private WindowSize As Integer, WS2 As Integer
Private BitLen As Integer
Private bitRevTable() As Single
Private normF As Single
Private Equalize() As Single
Private Envelope() As Single

Private FFTResult() As Single
Private FFTBuffer() As Single
Private cosLUT() As Single, sinLUT() As Single
Private FIRCoeffs() As Single
Private isEqualized As Boolean, hasEnvelope As Boolean

Private Const PI As Double = 3.14159265358979

Public Sub initFFTFilter(Window_Size As Integer)
WindowSize = Window_Size
WS2 = WindowSize

WS2 = WS2 \ 2 ' WS2 >>= 1
BitLen = CInt(Math.Log(WindowSize) / 0.693147180559945 + 0.5)
normF = 2 / Window_Size
hasEnvelope = False
isEqualized = False

Call initFFTtables

End Sub

Private Sub initFFTtables()

ReDim cosLUT(BitLen) As Single
ReDim sinLUT(BitLen) As Single

ReDim FFTBuffer(WindowSize, 2) As Single
ReDim FFTResult(WS2) As Single

Dim i As Integer
Dim Theta As Double
Theta = PI

'calculate sin/cos at BitLen angles
For i = 0 To BitLen
cosLUT(i) = Cos(Theta)
sinLUT(i) = Sin(Theta)
Theta = Theta * 0.5
Next

' precalculate bit reversal table

Dim j As Integer, bitm As Integer, temp As Integer
ReDim bitRevTable(WindowSize) As Single

For i = 0 To WindowSize
 bitRevTable(i) = i
Next

For i = 0 To WindowSize
 If (j > i) Then
  temp = bitRevTable(i)
  bitRevTable(i) = bitRevTable(j)
  bitRevTable(j) = temp
 End If
  
   bitm = WS2
    While (bitm >= 1 And j >= bitm)
      j = j - bitm
      bitm = bitm \ 2
      Wend
     j = j + bitm
   Next i
   
End Sub

' adapted from nullsoft
' reduces impact of bassy frequencies and slightly amplifies top range
Public Sub useEqualiser(Use As Boolean)
isEqualized = Use
 If (Use) Then
  Dim i As Integer
  Dim scaling As Single, invHalfnFreq As Single
  scaling = -0.02
  invHalfnFreq = 1# / WS2
  ReDim Equalize(WS2) As Single
   For i = 0 To WS2 - 1
    Equalize(i) = scaling * Math.Log((WS2 - i) * invHalfnFreq)
   Next
 End If
End Sub

' bell filter envelope to reduce artefacts caused by edges of standard filler rect.
' 0.0 < power < 2.0
Public Sub useEnvelope(Use As Boolean, Power As Single)
hasEnvelope = Use
 If (Use) Then
   Dim i As Integer
   Dim mult As Single
   
     mult = 1# / (WindowSize * 2 * PI)
     ReDim Envelope(WindowSize) As Single
     
      If (Power = 1#) Then
        For i = 0 To WindowSize
          Envelope(i) = 0.5 + 0.5 * Sin(i * mult - (PI / 2))
        Next i
      Else
        For i = 0 To WindowSize
         Envelope(i) = (0.5 + 0.5 * Sin(i * mult - (PI / 2))) ^ Power
        Next i
      End If
      
 End If
End Sub

' compute actual FFT with settings (eq,filter)
Public Sub computeFFT(waveInData() As Integer, FFTOutData() As Single)

Static u_r As Single, u_i As Single, w_r As Single, w_i As Single, t_r As Single, t_i As Single
Static l As Integer, le As Integer, le2 As Integer, j As Integer, jj As Integer, ip As Integer, i As Integer, ii As Integer, phi As Integer
Static idx As Integer, currFFT() As Single, t As Integer
Static ip1 As Single
 If (hasEnvelope) Then
   For i = 0 To WindowSize
    idx = bitRevTable(i)
    If (idx < WindowSize) Then
     FFTBuffer(i, 0) = waveInData(idx) * Envelope(idx)
    Else
     FFTBuffer(i, 0) = 0
    End If
     FFTBuffer(i, 1) = 0
   Next i
 Else
  For i = 0 To WindowSize
  idx = bitRevTable(i)
   If (idx < WindowSize) Then
    FFTBuffer(i, 0) = waveInData(idx)
   Else
    FFTBuffer(i, 0) = 0
   End If
    FFTBuffer(i, 1) = 0
   Next i
 End If
    
 l = 1
 le = 2
 phi = 0
 
 For l = 1 To BitLen
  le2 = le2 \ 2
  w_r = cosLUT(phi)
  w_i = sinLUT(phi)
  
  u_r = 1#
  u_i = 0#
  
    For j = 1 To le2
     For i = j To WindowSize Step le
     ip = i + le2
     ip1 = ip1 - 1
     ii = i - 1
     ReDim currFFT(idx) As Single
      
      For t = 0 To idx
       currFFT(t) = FFTBuffer(t, 0)
      Next t
      
     t_r = currFFT(0) * u_r - u_i * currFFT(1)
     t_i = currFFT(1) * u_r + u_i * currFFT(0)
     
      currFFT(0) = FFTBuffer(ii, 0) - t_r
      currFFT(1) = FFTBuffer(ii, 1) - t_i
      
      FFTBuffer(ii, 0) = FFTBuffer(ii, 0) + t_r
      FFTBuffer(ii, 1) = FFTBuffer(ii, 1) + t_i
     Next
    t_r = u_r * w_r - w_i * u_i
    u_i = w_r * u_i + w_i * u_r
    u_r = t_r
  Next
  le = le * 2
Next

' normalize bands or apply EQ
'Static currBin() As Single
'If (isEqualized) Then
' For i = 0 To WS2
' ReDim currBin(i + 1) As Single
'  For t = 0 To i
'   currBin(t) = FFTBuffer(t)
'   FFTResult(i) = Equalize(i) * Sqr(currBin(0) * currBin(0) + currBin(1) * currBin(1))
'  Next t
' Next i
' Else
' For i = 0 To WS2
'  ReDim currBin(0 To i + 1) As Single
'   For t = 0 To i
'    currBin(t) = FFTBuffer(t)
'   Next t
'  FFTResult(i) = normF * Sqr(currBin(0) * currBin(0) + currBin(1) * currBin(1))
' Next i
'End If
For t = 0 To UBound(FFTResult)
 FFTOutData(t) = FFTResult(t)
Next t
End Sub
