Attribute VB_Name = "modSpectrum"
'
' module containing API declarations for use in visualizations
'

Option Explicit

Public Const WAVE_FORMAT_PCM = 1
Public Const WHDR_DONE = &H1&              ' Done bit
Public Const WHDR_PREPARED = &H2&          ' Set if this header has been prepared
Public Const WHDR_BEGINLOOP = &H4&         ' Loop start block
Public Const WHDR_ENDLOOP = &H8&           ' Loop end block
Public Const WHDR_INQUEUE = &H10&          ' Reserved for driver
Public Const WIM_OPEN = &H3BE
Public Const WIM_CLOSE = &H3BF
Public Const WIM_DATA = &H3C0
Public Const ANGLENUMERATOR = 6.283185     ' 2 * Pi
Public Const NUMSAMPLES = 1024             ' Number of Samples
Public Const NUMBITS = 10                  ' Number of Bits

Public DevHandle As Long                    ' Handle of the open audio device
Public Visualizing As Boolean               ' Equalizer is working
Public Divisor As Long                      ' scale of the vis.
Public ScopeHeight As Long                  ' Saves time because hitting up a Long is faster
Public ReversedBits(0 To NUMSAMPLES - 1) As Long ' Bit reservation


'general
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long

'wave releated
Public Declare Function waveInAddBuffer Lib "winmm" (ByVal InputDeviceHandle As Long, ByVal WaveHdrPointer As Long, ByVal WaveHdrStructSize As Long) As Long
Public Declare Function waveInPrepareHeader Lib "winmm" (ByVal InputDeviceHandle As Long, ByVal WaveHdrPointer As Long, ByVal WaveHdrStructSize As Long) As Long
Public Declare Function waveInUnprepareHeader Lib "winmm" (ByVal InputDeviceHandle As Long, ByVal WaveHdrPointer As Long, ByVal WaveHdrStructSize As Long) As Long
Public Declare Function waveInGetNumDevs Lib "winmm" () As Long
Public Declare Function waveInGetDevCaps Lib "winmm" Alias "waveInGetDevCapsA" (ByVal uDeviceID As Long, ByVal WaveInCapsPointer As Long, ByVal WaveInCapsStructSize As Long) As Long
Public Declare Function waveInOpen Lib "winmm" (WaveDeviceInputHandle As Long, ByVal WhichDevice As Long, ByVal WaveFormatExPointer As Long, ByVal CallBack As Long, ByVal CallBackInstance As Long, ByVal Flags As Long) As Long
Public Declare Function waveInClose Lib "winmm" (ByVal WaveDeviceInputHandle As Long) As Long
Public Declare Function waveInStart Lib "winmm" (ByVal WaveDeviceInputHandle As Long) As Long
Public Declare Function waveInReset Lib "winmm" (ByVal WaveDeviceInputHandle As Long) As Long
Public Declare Function waveInStop Lib "winmm" (ByVal WaveDeviceInputHandle As Long) As Long

' wave format type
Public Type WaveFormatEx
    FormatTag As Integer
    Channels As Integer
    SamplesPerSec As Long
    AvgBytesPerSec As Long
    BlockAlign As Integer
    BitsPerSample As Integer
    ExtraDataSize As Integer
End Type

Public Type WAVEHDR
    lpData As Long
    dwBufferLength As Long
    dwBytesRecorded As Long
    dwUser As Long
    dwFlags As Long
    dwLoops As Long
    lpNext As Long
    Reserved As Long
End Type

Public Type WAVEINCAPS
    ManufacturerID As Integer
    ProductID As Integer
    DriverVersion As Long
    ProductName(1 To 32) As Byte
    Formats As Long
    Channels As Integer
    Reserved As Integer
End Type

' based on D.Cross's FFT code written in C

' Reverse array
Public Sub DoReverse()

    Dim i As Long
    For i = LBound(ReversedBits) To UBound(ReversedBits)
        ReversedBits(i) = ReverseBits(i, NUMBITS)
    Next
End Sub

' Reverse Bits
' just like nullsoft
Public Function ReverseBits(ByVal Index As Long, NUMBITS As Byte) As Long
    Dim i As Byte, Rev As Long
    For i = 0 To NUMBITS - 1
        Rev = (Rev * 2) Or (Index And 1)
        Index = Index \ 2
    Next
    ReverseBits = Rev
End Function

' Fast Fourier Tansform: FFT
Public Sub FFTAudio(RealIn() As Integer, RealOut() As Single)

    Static ImagOut(0 To NUMSAMPLES - 1) As Single
    Static i As Long, j As Long, k As Long, N As Long, BlockSize As Long, BlockEnd As Long
    Static DeltaAngle As Single, DeltaAr As Single
    Static Alpha As Single, Beta As Single
    Static TR As Single, TI As Single, AR As Single, AI As Single
    For i = 0 To (NUMSAMPLES - 1)
        j = ReversedBits(i)
        RealOut(j) = RealIn(i)
        ImagOut(j) = 0
    Next
    BlockEnd = 1
    BlockSize = 2
    Do While BlockSize <= NUMSAMPLES
        DeltaAngle = ANGLENUMERATOR / BlockSize
        Alpha = Sin(0.5 * DeltaAngle)
        Alpha = 2! * Alpha * Alpha
        Beta = Sin(DeltaAngle)
        i = 0
        Do While i < NUMSAMPLES
            AR = 1!
            AI = 0!
            j = i
            For N = 0 To BlockEnd - 1
                k = j + BlockEnd
                TR = AR * RealOut(k) - AI * ImagOut(k)
                TI = AI * RealOut(k) + AR * ImagOut(k)
                RealOut(k) = RealOut(j) - TR
                ImagOut(k) = ImagOut(j) - TI
                RealOut(j) = RealOut(j) + TR
                ImagOut(j) = ImagOut(j) + TI
                DeltaAr = Alpha * AR + Beta * AI
                AI = AI - (Alpha * AI - Beta * AR)
                AR = AR - DeltaAr
                j = j + 1
            Next N
            i = i + BlockSize
        Loop
        BlockEnd = BlockSize
        BlockSize = BlockSize * 2
    Loop
    
 Equalize RealOut

End Sub

' my own code

'equalize: reduces impact of bassy frequencies
'and slightly amplifies higher frequencies

Public Sub Equalize(InData() As Single)

 On Error Resume Next
  Dim i As Integer, Temp As Single
  

   For i = 0 To UBound(InData)
   'scaling = -0.01
    Temp = -0.01 * Math.Log(i) ' bassy frequcies are found in the LBound of the array
    InData(i) = InData(i) * Temp
   Next
 
End Sub

' Stop Greq
Public Sub DoStop()
    Call waveInReset(DevHandle)
    Call waveInClose(DevHandle)
    DevHandle = 0
End Sub

