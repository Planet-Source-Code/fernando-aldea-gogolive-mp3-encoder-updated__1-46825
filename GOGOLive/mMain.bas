Attribute VB_Name = "mMain"
''''''''''''''''''''''''''''''''''''''''''''''
''    Module written by Fernando Aldea G.   ''
''    e-mail: fernando_aldea@terra.cl       ''
''    web: orbita.starmedia.com/gogolive/   ''
''    Release Juny, 2003                    ''
''                                          ''
''   sorry for not translate this completly ''
''    & sorry about my English!             ''
''''''''''''''''''''''''''''''''''''''''''''''

Public Const SRCCOPY = &HCC0020 ' (DWORD) dest = source


'kernel
Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long

Public mPipe As New cPipe

'Public WAV_Frequency As Long
'Public WAV_Bits As Long
'Public WAV_Stereo As Boolean

Public MP3_Frequency As Long
Public MP3_Kbps As Long
Public MP3_Mode As Long
Public MP3_Enphasis As Long
Public MP3_outFile As String
Public MP3_CPU As Boolean
Public MP3_PSY As Boolean
Public MP3_MMX As Boolean
Public MP3_LPF16 As Boolean

Public TimerIni As Single
Public nBytesPCMacum As Long
Public Recording As Boolean
Public Monitoring As Boolean
Public BUFFER_LENGTH As Long
Public OUTFILE_DEFAULT As String
'Public SyncCD As Boolean


Sub Main()
    MP3_Frequency = 44100
    MP3_Kbps = 128
    MP3_Mode = 2 'joint stereo
    MP3_Enphasis = 0
    MP3_PSY = True
    MP3_CPU = True
    MP3_LPF16 = False
    MP3_MMX = True
    OUTFILE_DEFAULT = "c:\newGOGO.mp3"
    MP3_outFile = OUTFILE_DEFAULT
    BUFFER_LENGTH = 10000000 '15000000
    
    frmGogoLive.Show
    
            
End Sub


Sub StartRecord()
    Dim tStereo As Boolean
    
    If MP3_Mode <> MC_MODE_MONO Then
        tStereo = True
    Else
        tStereo = False
    End If
    
    'Initialize Wave Service
    If Not mWaveCard.Initialize(MP3_Frequency, tStereo) Then MsgBox "Error in WAVE", vbCritical
    'Initialize Gogo Library
    If Not InitializeGOGO(MP3_outFile, MP3_Kbps, MP3_Mode, MP3_CPU, MP3_MMX, MP3_PSY, MP3_LPF16) Then MsgBox "Error in GOGO", vbCritical
    'Show parameters GOGO library
    ShowGOGOInfo
        
    TimerIni = Timer  'security for timer record
    nBytesPCMacum = -1
    
    'Start receiving audio input
    If StartInput() Then
        Recording = True
        StartEncode
    Else
        MsgBox "Error in StartInput", vbCritical
    End If
    
    'Release Gogo Library
    Finalizar
   
End Sub

Sub EndRecord()
    StopInput
    Recording = False
End Sub

Sub ShowGOGOInfo()
    Dim mStr As String
    
    'If InfoGogo(MG_USEPSY) = 1 Then frmGogoLive.lblIzq(0).Enabled = False
    'If InfoGogo(MG_USEMMX) = 1 Then frmGogoLive.lblIzq(2).Enabled = False
    
    frmGogoLive.lblDer(0).Caption = InfoGogo(MG_OUTFREQ) & " Hz": frmGogoLive.lblDer(0).Visible = True
    frmGogoLive.lblDer(1).Caption = InfoGogo(MG_BITRATE) & " Kbps": frmGogoLive.lblDer(1).Visible = True
            
    Call Opciones_Modo_GOGO(InfoGogo(MC_ENCODEMODE), mStr)
    frmGogoLive.lblDer(2).Caption = mStr: frmGogoLive.lblDer(2).Visible = True
        
    If InfoGogo(MG_USEPSY) Then frmGogoLive.lblIzq(0).Enabled = False: frmGogoLive.lblIzq(0).Visible = True
    If InfoGogo(MG_USEMMX) Then frmGogoLive.lblIzq(2).Enabled = False: frmGogoLive.lblIzq(2).Visible = True
    If MP3_CPU Then frmGogoLive.lblIzq(3).Visible = True: frmGogoLive.lblIzq(3).Enabled = False
    If MP3_LPF16 Then frmGogoLive.lblIzq(1).Visible = True: frmGogoLive.lblIzq(1).Enabled = False
    
    'frmGogoLive.lblSong.Caption = InfoGogo(MG_OUTPUTFILE)
    frmGogoLive.lblSong.Caption = MP3_outFile
            
    
    
End Sub

Public Function PauseRecord()
    PauseInput
End Function


'util
Public Function GetAddressofFunction(ByVal pFunction As Long) As Long
    GetAddressofFunction = pFunction
End Function
