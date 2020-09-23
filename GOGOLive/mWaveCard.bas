Attribute VB_Name = "mWaveCard"
''''''''''''''''''''''''''''''''''''''''''''''
''    Module adapted by Fernando Aldea G.   ''
''    e-mail: fernando_aldea@terra.cl       ''
''    web: orbita.starmedia.com/gogolive/   ''
''    Release Juny, 2003                    ''
''                                          ''
''    sorry for not translate this completly''
''    & sorry about my English!             ''
''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Public Const CALLBACK_FUNCTION = &H30000
Public Const CALLBACK_WINDOW = &H10000      '  dwCallback is a HWND
Public Const MM_WIM_DATA = &H3C0
Public Const WHDR_DONE = &H1         '  done bit
Public Const WIM_DATA = MM_WIM_DATA
Public Const GMEM_FIXED = &H0         ' Global Memory Flag used by GlobalAlloc functin
Public Const NUM_BUFFERS = 10
Public BUFFER_SIZE As Long  '= 8192
Public Const DEVICEID = -1
Public Const GWL_WNDPROC = -4

'callback mode
Private Const CM_WINDOWS = 1
Private Const CM_FUNCTION = 2



Type WAVEHDR
   lpData As Long          ' Address of the waveform buffer.
   dwBufferLength As Long  ' Length, in bytes, of the buffer.
   dwBytesRecorded As Long ' When the header is used in input, this member specifies how much
                           ' data is in the buffer.

   dwUser As Long          ' User data.
   dwFlags As Long         ' Flags supplying information about the buffer. Set equal to zero.
   dwLoops As Long         ' Number of times to play the loop. Set equal to zero.
   lpNext As Long          ' Not used
   reserved As Long        ' Not used
End Type

Type WAVEFORMAT
   wFormatTag As Integer
   nChannels As Integer
   nSamplesPerSec As Long
   nAvgBytesPerSec As Long
   nBlockAlign As Integer
   wBitsPerSample As Integer
   cbSize As Integer
End Type

Declare Function waveInOpen Lib "winmm.dll" (lphWaveIn As Long, ByVal uDeviceID As Long, lpFormat As WAVEFORMAT, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Declare Function waveInPrepareHeader Lib "winmm.dll" (ByVal hWaveIn As Long, lpWaveInHdr As WAVEHDR, ByVal uSize As Long) As Long
Declare Function waveInReset Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Declare Function waveInStart Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Declare Function waveInStop Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Declare Function waveInUnprepareHeader Lib "winmm.dll" (ByVal hWaveIn As Long, lpWaveInHdr As WAVEHDR, ByVal uSize As Long) As Long
Declare Function waveInClose Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Declare Function waveInGetErrorText Lib "winmm.dll" Alias "waveInGetErrorTextA" (ByVal err As Long, ByVal lpText As String, ByVal uSize As Long) As Long
Declare Function waveInAddBuffer Lib "winmm.dll" (ByVal hWaveIn As Long, lpWaveInHdr As WAVEHDR, ByVal uSize As Long) As Long

' Global Memory Flags
'Public Const GMEM_FIXED = &H0
Public Const GMEM_ZEROINIT = &H40
Public Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)

Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Declare Function GlobalLock Lib "kernel32" (ByVal hmem As Long) As Long
Declare Function GlobalFree Lib "kernel32" (ByVal hmem As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Sub CopyStringFromStruct Lib "kernel32" Alias "RtlMoveMemory" (ByVal a As String, p As Any, ByVal cb As Long)
Public Declare Sub CopyStructFromPtr Lib "kernel32" Alias "RtlMoveMemory" (struct As Any, ByVal ptr As Long, ByVal cb As Long)
Public Declare Sub CopyPtrFromStruct Lib "kernel32" Alias "RtlMoveMemory" (ByVal ptr As Long, struct As Any, ByVal cb As Long)
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByRef lParam As WAVEHDR) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Dim hWaveIn As Long
Dim wformat As WAVEFORMAT
Dim hmem(NUM_BUFFERS) As Long
Dim inHdr(NUM_BUFFERS) As WAVEHDR
Dim fRecording As Boolean
Dim fPausing As Boolean
Dim lpPrevWndProc As Long
Dim CallbackMode As Integer
Public largoSamples As Long
Dim msg As String * 200
Dim hWnd As Long

Public Property Get isPausing() As Boolean
    isPausing = fPausing
End Property

Public Property Get isRecording() As Boolean
    isRecording = fRecording
End Property

Public Property Get nAvgBytesPerSec() As Long
    nAvgBytesPerSec = wformat.nAvgBytesPerSec
End Property

'start audio input from soundcard
Function StartInput(Optional lBuffer As Long) As Boolean
    Dim i As Long
    Dim rc As Long

    If CallbackMode <= 0 Then
        MsgBox "Not initialize", vbCritical
        Exit Function
    End If
    
    If fRecording Then
        StartInput = True
        Exit Function
    End If
    
    BUFFER_SIZE = (wformat.nSamplesPerSec * wformat.nBlockAlign * wformat.nChannels * 0.1) - ((wformat.nSamplesPerSec * wformat.nBlockAlign * wformat.nChannels * 0.1) Mod (wformat.nBlockAlign))
    BUFFER_SIZE = BUFFER_SIZE * 2
    
    If lBuffer > 0 Then BUFFER_SIZE = lBuffer
    For i = 0 To NUM_BUFFERS - 1
        hmem(i) = GlobalAlloc(&H40, BUFFER_SIZE)
        inHdr(i).lpData = GlobalLock(hmem(i))
        inHdr(i).dwBufferLength = BUFFER_SIZE
        inHdr(i).dwFlags = 0
        inHdr(i).dwLoops = 0
    Next
    
    
    If CallbackMode = CM_FUNCTION Then rc = waveInOpen(hWaveIn, DEVICEID, wformat, AddressOf WaveProc, 0, CALLBACK_FUNCTION)
    If CallbackMode = CM_WINDOWS Then rc = waveInOpen(hWaveIn, DEVICEID, wformat, AddressOf WaveProc, 0, CALLBACK_FUNCTION)
        
    If rc <> 0 Then
        waveInGetErrorText rc, msg, Len(msg)
        MsgBox msg, vbCritical
        StartInput = False
        Exit Function
    End If

    For i = 0 To NUM_BUFFERS - 1
        rc = waveInPrepareHeader(hWaveIn, inHdr(i), Len(inHdr(i)))
        If (rc <> 0) Then
            waveInGetErrorText rc, msg, Len(msg)
            MsgBox msg, vbCritical
        End If
    Next

    For i = 0 To NUM_BUFFERS - 1
        addData inHdr(i)
        If (rc <> 0) Then
            waveInGetErrorText rc, msg, Len(msg)
            MsgBox msg, vbCritical
        End If
    Next

    largoSamples = (wformat.wBitsPerSample / 8) * wformat.nChannels
    If largoSamples = 0 Then Beep: largoSamples = 1
    
    fRecording = True
    rc = waveInStart(hWaveIn)
    StartInput = True
    
End Function

Sub addData(iHdr As WAVEHDR)
    Dim sBuff  As String
    Dim rc As Long
    
    rc = waveInAddBuffer(hWaveIn, iHdr, Len(iHdr))
    sBuff = Space(BUFFER_SIZE)
    CopyMemory ByVal sBuff, ByVal iHdr.lpData, BUFFER_SIZE
End Sub

' Stop receiving audio input on the soundcard
Sub StopInput()
    Dim iRet As Long
    Dim i As Long
    
    fRecording = False
    iRet = waveInReset(hWaveIn)
    iRet = waveInStop(hWaveIn)
    For i = 0 To NUM_BUFFERS - 1
        waveInUnprepareHeader hWaveIn, inHdr(i), Len(inHdr(i))
        GlobalFree hmem(i)
    Next
    iRet = waveInClose(hWaveIn)
End Sub

' Initialize input soundcard service by windows callback mode
Public Function Initialize(ByVal Frequency As Long, ByVal Stereo As Boolean, Optional hwndIn As Long) As Boolean
   
    'Set The WAV wformat
    wformat.wFormatTag = 1
    If Stereo Then
        wformat.nChannels = 2
    Else
        wformat.nChannels = 1
    End If
    wformat.wBitsPerSample = 16
    wformat.nSamplesPerSec = Frequency
    wformat.nBlockAlign = wformat.nChannels * wformat.wBitsPerSample / 8
    wformat.nAvgBytesPerSec = wformat.nSamplesPerSec * wformat.nBlockAlign
    wformat.cbSize = Len(wformat)
    
    If hwndIn = 0 Then
        CallbackMode = CM_FUNCTION
    Else
        hWnd = hwndIn
        lpPrevWndProc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WindowProc)
        CallbackMode = CM_WINDOWS
    End If
    Initialize = True
End Function

'procedure for windows callback mode
Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByRef wavhdr As WAVEHDR) As Long
    Dim i As Integer
    Dim rc As Long
    
    If uMsg = WIM_DATA Then
        ' Process sound buffer if recording
        If (fRecording) Then
            For i = 0 To (NUM_BUFFERS - 1)
                If inHdr(i).dwFlags And WHDR_DONE Then
                    rc = waveInAddBuffer(hWaveIn, inHdr(i), Len(inHdr(i)))
                    If rc <> 0 Then
                        MsgBox "Failed (WaveInAddBuffer)", vbCritical
                    End If
                    If Not (fPausing) Then
                        frmGogoLive.callBackWave inHdr(i).lpData, BUFFER_SIZE   '<--------  sub necesaria en el form
                    End If
                End If
            Next i
        End If
    End If
    WindowProc = CallWindowProc(lpPrevWndProc, hw, uMsg, wParam, wavhdr)
End Function

'procedure for function callback mode
Sub WaveProc(ByVal hw As Long, ByVal uMsg As Long, ByVal dwInstance As Long, ByVal wParam As Long, ByRef wavhdr As WAVEHDR)
    Dim i As Integer
    Dim rc As Long
    
    ' Process sound buffer if recording
    If (fRecording) Then
        For i = 0 To (NUM_BUFFERS - 1)
            If inHdr(i).dwFlags And WHDR_DONE Then
                rc = waveInAddBuffer(hWaveIn, inHdr(i), Len(inHdr(i)))
                If rc <> 0 Then
                    MsgBox "Failed (waveInAddBuffer)", vbCritical
                End If
                If Not (fPausing) Then
                    frmGogoLive.callBackWave inHdr(i).lpData, BUFFER_SIZE   '<--------  sub necesaria en el form
                End If
            End If
        Next i
   End If
End Sub

' Stop receiving audio input on the soundcard
Sub PauseInput()
    fPausing = Not fPausing
End Sub

Function iniMonitor() As Boolean
    'Set The WAV wformat
    wformat.wFormatTag = 1
    If MP3_Mode <> MC_MODE_MONO Then
        wformat.nChannels = 2
    Else
        wformat.nChannels = 1
    End If
    
    wformat.wBitsPerSample = 16
    wformat.nSamplesPerSec = 22050 'MP3_Frequency
    wformat.nBlockAlign = wformat.nChannels * wformat.wBitsPerSample / 8
    wformat.nAvgBytesPerSec = wformat.nSamplesPerSec * wformat.nBlockAlign
    wformat.cbSize = Len(wformat)

    CallbackMode = CM_FUNCTION
    If StartInput(wformat.nAvgBytesPerSec / 10) Then iniMonitor = True
    
End Function
Sub TerminarMonitor()
    StopInput
End Sub
