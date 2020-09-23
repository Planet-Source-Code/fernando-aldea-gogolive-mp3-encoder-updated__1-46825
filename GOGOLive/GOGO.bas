Attribute VB_Name = "GOGO"
''''''''''''''''''''''''''''''''''''''''''''''
''    Module adapted by Fernando Aldea G.   ''
''    e-mail: fernando_aldea@terra.cl       ''
''    web: orbita.starmedia.com/gogolive/   ''
''    Release Juny, 2003                    ''
''                                          ''
''    sorry about my English!               ''
''''''''''''''''''''''''''''''''''''''''''''''


' Api. Global Memory Flags
  Private Const GMEM_FIXED = &H0
Private Const GMEM_ZEROINIT = &H40
Private Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)

Public Const ULONG_MAX = 255


' Configuration
Public Const MC_INPUTFILE = 1
Public Const MC_INPDEV_FILE = 0              ' input device is file;ì¸óÕÉfÉoÉCÉXÇÕÉtÉ@ÉCÉã
Public Const MC_INPDEV_STDIO = 1             '                 stdin;ì¸óÕÉfÉoÉCÉXÇÕïWèÄì¸óÕ
Public Const MC_INPDEV_USERFUNC = 2          '       defined by user;ì¸óÕÉfÉoÉCÉXÇÕÉÜÅ[ÉUÅ[íËã`
Public Const MC_OUTPUTFILE = 2
    Public Const MC_OUTDEV_FILE = 0              ' output device is file
    Public Const MC_OUTDEV_STDOUT = 1            '                  stdout
    Public Const MC_OUTDEV_USERFUNC = 2          '        defined by user
    Public Const MC_OUTDEV_USERFUNC_WITHVBRTAG = 3       '       defined by user
Public Const MC_ENCODEMODE = 3
    Public Const MC_MODE_MONO = 0                ' mono;ÉÇÉmÉâÉã
    Public Const MC_MODE_STEREO = 1              ' stereo;ÉXÉeÉåÉI
    Public Const MC_MODE_JOINT = 2               ' joint-stereo;ÉWÉáÉCÉìÉg
    Public Const MC_MODE_MSSTEREO = 3            ' mid/side stereo;É~ÉbÉhÉTÉCÉh
    Public Const MC_MODE_DUALCHANNEL = 4         ' dual channel;ÉfÉÖÉAÉãÉ`ÉÉÉlÉã
Public Const MC_BITRATE = 4
Public Const MC_INPFREQ = 5
Public Const MC_OUTFREQ = 6
Public Const MC_STARTOFFSET = 7
Public Const MC_USEPSY = 8
Public Const MC_USELPF16 = 9
Public Const MC_USEMMX = 10                      ' MMX
Public Const MC_USE3DNOW = 11                    ' 3DNow!
Public Const MC_USEKNI = 12                      ' SSE=KNI
Public Const MC_USEE3DNOW = 13                   ' Enhanced 3D Now!
Public Const MC_USESPC1 = 14                     ' special switch for debug
Public Const MC_USESPC2 = 15                     ' special switch for debug
Public Const MC_ADDTAG = 16
Public Const MC_EMPHASIS = 17
Public Const MC_EMP_NONE = 0                 ' no empahsis
Public Const MC_EMP_5015MS = 1               ' 50/15ms
Public Const MC_EMP_CCITT = 3                ' CCITT
Public Const MC_VBR = 18
Public Const MC_CPU = 19
Public Const MC_BYTE_SWAP = 20
Public Const MC_8BIT_PCM = 21
Public Const MC_MONO_PCM = 22
Public Const MC_TOWNS_SND = 23
Public Const MC_THREAD_PRIORITY = 24
Public Const MC_READTHREAD_PRIORITY = 25
Public Const MC_OUTPUT_FORMAT = 26
Public Const MC_OUTPUT_NORMAL = 0            ' mp3+TAG=see MC_ADDTAG
Public Const MC_OUTPUT_RIFF_WAVE = 1         ' RIFF/WAVE
Public Const MC_OUTPUT_RIFF_RMP = 2          ' RIFF/RMP
Public Const MC_RIFF_INFO = 27
Public Const MC_VERIFY = 28
Public Const MC_OUTPUTDIR = 29
Public Const MC_VBRBITRATE = 30
Public Const MC_ENHANCEDFILTER = 31
Public Const MC_MSTHRESHOLD = 32

'Language
Public Const MC_LANG = 33
Public Const MC_MAXFILELENGTH = 34
Public Const MC_MAXFLEN_IGNORE = ULONG_MAX
Public Const MC_MAXFLEN_WAVEHEADER = ULONG_MAX - 1
Public Const MC_OUTSTREAM_BUFFERD = 35
Public Const MC_OBUFFER_ENABLE = 1
Public Const MC_OBUFFER_DISABLE = 0

'Errors
Public Const ME_NOERR = 0                        ' return normally
Public Const ME_EMPTYSTREAM = 1                  ' stream becomes empty
Public Const ME_HALTED = 2                       ' stopped by user
Public Const ME_MOREDATA = 3
Public Const ME_INTERNALERROR = 10               ' internal error;
Public Const ME_PARAMERROR = 11                  ' parameters error;
Public Const ME_NOFPU = 12                       ' no FPU;
Public Const ME_INFILE_NOFOUND = 13              ' can't open input file;
Public Const ME_OUTFILE_NOFOUND = 14             ' can't open output file;
Public Const ME_FREQERROR = 15                   ' frequency is not good
Public Const ME_BITRATEERROR = 16                ' bitrate is not good
Public Const ME_WAVETYPE_ERR = 17                ' WAV format is not good
Public Const ME_CANNOT_SEEK = 18                 ' can't seek
Public Const ME_BITRATE_ERR = 19                 ' only for compatibility
Public Const ME_BADMODEORLAYER = 20              ' mode/layer not good
Public Const ME_NOMEMORY = 21                    ' fail to allocate memory
Public Const ME_CANNOT_SET_SCOPE = 22            ' thread error
Public Const ME_CANNOT_CREATE_THREAD = 23        ' fail to create thear
Public Const ME_WRITEERROR = 24                  ' lock of capacity of disk


' getting configuration
Public Const MG_INPUTFILE = 1                    ' name of input file
Public Const MG_OUTPUTFILE = 2                   ' name of output file
Public Const MG_ENCODEMODE = 3                   ' type of encoding
Public Const MG_BITRATE = 4                      ' bitrate
Public Const MG_INPFREQ = 5                      ' input frequency
Public Const MG_OUTFREQ = 6                      ' output frequency   ;≈o˘-ƒ≥ˆg…ˆ
Public Const MG_STARTOFFSET = 7                  ' offset of input PCM;‚X‚^¸[‚g‚I‚t‚Z‚b‚g
Public Const MG_USEPSY = 8                       ' psycho-acoustics   ;…S˘ÿÎ≠…-È≠ƒg˘pÈ¿È⁄/È¡È+ÈÛ
Public Const MG_USEMMX = 9                       ' MMX
Public Const MG_USE3DNOW = 10                    ' 3DNow!
Public Const MG_USEKNI = 11                      ' SSE=KNI
Public Const MG_USEE3DNOW = 12                   ' Enhanced 3DNow!

Public Const MG_USESPC1 = 13                     ' special switch for debug
Public Const MG_USESPC2 = 14                     ' special switch for debug
Public Const MG_COUNT_FRAME = 15                 ' amount of frame
Public Const MG_NUM_OF_SAMPLES = 16              ' number of sample for 1 frame;1‚t‚Ó¸[‚«È·È¢ÈﬁÈ¶‚T‚Ù‚v‚Ô…ˆ
Public Const MG_MPEG_VERSION = 17                ' MPEG VERSION
Public Const MG_READTHREAD_PRIORITY = 18         ' thread priority to read for BeOS




Enum t_lang
    tLANG_UNKNOWN
    tLANG_JAPANESE_SJIS
    tLANG_JAPANESE_EUC
    tLANG_ENGLISH
    tLANG_GERMAN
    tLANG_SPANISH
End Enum

Type MCP_INPDEV_USERFUNC
    pUserFunc As Long   ' pointer to user-function for call-back or MPGE_NULL_FUNC if none
    nSize As Long       ' size of file or MC_INPDEV_MEMORY_NOSIZE if unknown
    nBit As Long        ' nBit = 8 or 16
    nFreq As Long       'input frequency
    nChn As Long        'number of channel(1 or 2)
End Type


  Declare Function MPGE_closeCoderVB Lib "gogo.dll" () As Long
  Declare Function MPGE_detectConfigureVB Lib "gogo.dll" () As Long
  Declare Function MPGE_endCoderVB Lib "gogo.dll" () As Long
  Declare Function MPGE_getConfigureVB Lib "gogo.dll" (ByVal Mode As Long, para1 As Any) As Long
  Declare Function MPGE_getUnitStatesVB Lib "gogo.dll" (unit As Long) As Long
  Declare Function MPGE_getVersionVB Lib "gogo.dll" (pNum As Long, pStr As String) As Long
  Declare Function MPGE_initializeWorkVB Lib "gogo.dll" () As Long
  Declare Function MPGE_processFrameVB Lib "gogo.dll" () As Long
  Declare Function MPGE_setConfigureVB Lib "gogo.dll" (ByVal Mode As Long, ByVal dwPara1 As Long, ByVal dwPara2 As String) As Long
  Declare Function MPGE_setConfigureVB2 Lib "gogo.dll" Alias "MPGE_setConfigureVB" (ByVal Mode As Long, ByVal dwPara1 As Long, dwPara2 As MCP_INPDEV_USERFUNC) As Long
  Declare Function MPGE_setConfigureVB3 Lib "gogo.dll" Alias "MPGE_setConfigureVB" (ByVal Mode As Long, ByVal dwPara1 As Long, dwPara2 As Long) As Long
  Declare Function MPGE_processTrack Lib "gogo.dll" (ByRef frameNum As Integer) As Long
  Declare Function MPGE_processTrack2 Lib "gogo.dll" Alias "MPGE_processTrack" (ByVal frameNum As Long) As Long
 
 
 Global curFrame As Long


Public Function InitializeGOGO(ByVal outFile As String, ByVal Kbps As Long, ByVal Mode As Long, Optional CPU As Boolean, Optional MMX As Boolean, Optional PSY As Boolean, Optional LPF16 As Boolean) As Boolean
    Dim Resp As Long
    
    InitializeGOGO = False
    'Call MPGE_closeCoderVB

        ' 6.free gogod.ll
    'Call MPGE_endCoderVB
        
    Resp = MPGE_initializeWorkVB()
    Call MPGE_endCoderVB
    
    Resp = MPGE_setConfigureVB(MC_LANG, tLANG_SPANISH, 0)
    GoSub verify
    
    Resp = MPGE_setConfigureVB(MC_THREAD_PRIORITY, 3, 0)
    GoSub verify
        
    'resp = MPGE_setConfigureVB(MC_INPUTFILE, MC_INPDEV_FILE, "c:\music.wav")
    'GoSub verify
      
    Dim Func As MCP_INPDEV_USERFUNC
    Func.nBit = 16
    If MP3_Mode <> MC_MODE_MONO Then Func.nChn = 2 Else Func.nChn = 1
    Func.nFreq = MP3_Frequency
    Func.nSize = 500000
    'Func.pUserFunc = Func2Point(AddressOf laFuncion)
    Func.pUserFunc = GetAddressofFunction(AddressOf laFuncion)

    
    Resp = MPGE_setConfigureVB2(MC_INPUTFILE, MC_INPDEV_USERFUNC, Func)
    GoSub verify
        
        
        
    '---------- output ---------
    
    Resp = MPGE_setConfigureVB(MC_OUTPUTFILE, MC_OUTDEV_FILE, outFile)
    GoSub verify
    
    Resp = MPGE_setConfigureVB3(MC_BITRATE, Kbps, 0&)
    GoSub verify
    
    Resp = MPGE_setConfigureVB3(MC_ENCODEMODE, Mode, 0&)  'stereo
    GoSub verify
    
    If CPU Then
        Resp = MPGE_setConfigureVB3(MC_CPU, 1, 0&)
    Else
        Resp = MPGE_setConfigureVB3(MC_CPU, 0, 0&)
    End If
    GoSub verify
    
    If MMX Then
        Resp = MPGE_setConfigureVB3(MC_USEMMX, 1, 0&)
    Else
        Resp = MPGE_setConfigureVB3(MC_USEMMX, 0, 0&)
    End If
    GoSub verify
    
    If PSY Then
        Resp = MPGE_setConfigureVB3(MC_USEPSY, 1, 0&)
    Else
        Resp = MPGE_setConfigureVB3(MC_USEPSY, 0, 0&)
    End If
    GoSub verify
    
    If LPF16 Then
        Resp = MPGE_setConfigureVB3(MC_USELPF16, 1, 0&) ' Abs(CLng(MP3_LPF16))
    Else
        Resp = MPGE_setConfigureVB3(MC_USELPF16, 1, 0&) ' Abs(CLng(MP3_LPF16))
    End If
    GoSub verify
        
    Resp = MPGE_detectConfigureVB()
    GoSub verify
        
    InitializeGOGO = True
    
    Exit Function
    
verify:
    If Resp <> 0 Then
        MsgBox GetError(CInt(Resp))
        Exit Function
    End If
    Return
End Function


Function GetError(n As Long) As String

Select Case n
Case ME_NOERR
    GetError = " return normally"
Case ME_EMPTYSTREAM
    GetError = " stream becomes empty" '‚X‚g‚Ë¸[‚«È¨Ï+Ó“È+∆BÈ¡È¢
Case ME_HALTED
    GetError = " stopped by user" '=‚Â¸[‚U¸[È¶ƒﬁÈ+ÈµÈﬁ ∆Â∆fÈ¶È€È¢
Case ME_MOREDATA
Case ME_INTERNALERROR
    GetError = " internal error" ' Ù”Úˆ‚G‚Î¸[
Case ME_PARAMERROR
    GetError = " parameters error" '…¶∆ﬁÈ+‚p‚Î‚¸¸[‚^¸[‚G‚Î¸[
Case ME_NOFPU
    GetError = " no FPU" 'FPUÈ≠ÊÚ∆‡È¡È-ÈÛÈ+ÈÛ!!
Case ME_INFILE_NOFOUND
    GetError = " open input file"
Case ME_OUTFILE_NOFOUND
    GetError = " open output file" '≈o˘-‚t‚@‚C‚ÔÈ≠…¶È¡È°ËJÈªÈ+ÈÛ
Case ME_FREQERROR
    GetError = " frequency is not good" 'Ù≥≈o˘-ƒ≥ˆg…ˆÈ¨…¶È¡È°È+ÈÛ
Case ME_BITRATEERROR
    GetError = " bitrate is not good" '≈o˘-‚r‚b‚g‚Ó¸[‚gÈ¨…¶È¡È°È+ÈÛ
Case ME_WAVETYPE_ERR
    GetError = " WAV format is not good" '‚E‚F¸[‚u‚^‚C‚vÈ¨…¶È¡È°È+ÈÛ
Case ME_CANNOT_SEEK
    GetError = "  seek" '…¶È¡È°‚V¸[‚N≈o˘ÍÈ+ÈÛ
Case ME_BITRATE_ERR
    GetError = " only for compatibility" '‚r‚b‚g‚Ó¸[‚g…¶∆ﬁÈ¨…¶È¡È°È+ÈÛ
Case ME_BADMODEORLAYER
    GetError = " mode/layer not good" '‚È¸[‚h¸E‚Ó‚C‚‰È¶…¶∆ﬁÍ+≈›
Case ME_NOMEMORY
    GetError = " fail to allocate memory" '‚¸‚È‚Ë‚A‚Ï¸[‚P¸[‚V‚Á‚Ùƒ©ˆs
Case ME_CANNOT_SET_SCOPE
    GetError = " thread error" '‚X‚Ó‚b‚hÊ´…Ω‚G‚Î¸[=pthread only
Case ME_CANNOT_CREATE_THREAD
    GetError = " fail to create thear" '‚X‚Ó‚b‚h…¬…º‚G‚Î¸[
Case ME_WRITEERROR
    GetError = " lock of capacity of disk" 'ÔLÎªˆ}Ê¶È¶˘e˘-ÚsÊΩ
End Select
GetError = GetError & "(" & n & ")"
End Function


Function laFuncion(ByVal hBuf As Long, ByVal Largo As Long) As Long
            
    laFuncion = frmGogoLive.callBackMp3(hBuf, Largo)
    
End Function

Function StartEncode() As Long
        'Dim hFile As Long
        Dim rval As Long
        Dim Ptr1 As Long
        
        Ptr1 = GlobalAlloc(GPTR, BLOQUE)
        
        ' get the amount of frames
        Dim TotalFrames As Long
        'Dim nFrames As Integer
                
        rval = MPGE_getConfigureVB(MG_COUNT_FRAME, TotalFrames)
        If rval <> 0 Then GoTo ssalir
        
        curFrame = 0
                
        ' start to encode
        Do
                rval = MPGE_processFrameVB()
                If rval = ME_NOERR Then curFrame = curFrame + 1
                DoEvents
        Loop Until (rval <> ME_NOERR And rval <> ME_MOREDATA)
        
        
        If rval <> ME_EMPTYSTREAM Then
            MsgBox ("ERROR: errcode = " & rval & Chr$(13) & GetError(rval))
        End If
ssalir:
        ' 5. end of encoding
        Call MPGE_closeCoderVB

        ' 6.free gogod.ll
        Call MPGE_endCoderVB
               
               
        StartEncode = 1
        
End Function

Public Sub Info(mp3File As String, mp3Frec As Long, Stereo As Boolean, mp3Bitrate As Long)
    Dim pun As Long
    
    mp3File = "                          "
    pun = GlobalAlloc(GPTR, Len(mp3File))
    
    
    'obtener archivo out
    Resp = MPGE_getConfigureVB(MG_OUTPUTFILE, ByVal pun)
    GoSub verify
    CopyMemory mp3File, pun, Len(mp3File)
    
    Dim valor As Long
    
    'Resp = MPGE_getConfigureVB(MG_INPFREQ, valor)
    'frmGogoLive.ListMp3.AddItem "frec IN : " & valor
    Resp = MPGE_getConfigureVB(MG_OUTFREQ, valor)
    mp3Frec = valor
    Resp = MPGE_getConfigureVB(MG_ENCODEMODE, valor)
    If valor = 2 Then Stereo = True
    Resp = MPGE_getConfigureVB(MG_BITRATE, valor)
    mp3Bitrate = valor
    'Resp = MPGE_getConfigureVB(MG_MPEG_VERSION, valor)
    'frmGogoLive.ListMp3.AddItem "mpeg ver : " & valor
    'Resp = MPGE_getConfigureVB(MG_COUNT_FRAME, valor)
    'frmGogoLive.ListMp3.AddItem "c frames : " & valor
    'Resp = MPGE_getConfigureVB(MG_STARTOFFSET, valor)
    'frmGogoLive.ListMp3.AddItem "offset input PCM: " & valor
    'Resp = MPGE_getConfigureVB(MG_NUM_OF_SAMPLES, valor)
    'frmGogoLive.ListMp3.AddItem "n samples x frame IN : " & valor
    
    GlobalFree pun
    
    Exit Sub
verify:
    If Resp <> 0 Then
        MsgBox GetError(CInt(Resp))
        GlobalFree pun
        Exit Sub
    End If
    Return

End Sub

Function Version() As String
    Dim mNum As Long
    Dim mStr As String
    Dim Resp As Long
    
    mNum = GlobalAlloc(GPTR, 20) 'PUNTERO
    mStr = Space$(255)
    
    Resp = MPGE_getVersionVB(ByVal mNum, ByVal mStr)
    If Resp = 0 Then
        Version = mStr
    Else
        Version = GetError(Resp)
    End If
    
    GlobalFree mNum 'libero puntero
    
End Function


Function Finalizar()
    ' 5. end of encoding
        Call MPGE_closeCoderVB

        ' 6.free gogod.ll
        Call MPGE_endCoderVB
        
End Function



Function Value2String(ByVal Mode As Long, ByVal Param As Long) As String
    
Select Case Mode
 
Case MC_ENCODEMODE
    Select Case Param
        Case MC_MODE_MONO
            Value2String = "mono"
        Case MC_MODE_STEREO
            Value2String = "stereo"
        Case MC_MODE_JOINT
            Value2String = "joint"
        Case MC_MODE_MSSTEREO
            Value2String = "msstereo"
        Case MC_MODE_DUALCHANNEL
            Value2String = "dual"
    End Select
Case MC_EMPHASIS ' 17
    Select Case Param
        Case MC_EMP_NONE
            Value2String = "NO"
        Case MC_EMP_5015MS
            Value2String = "50/15ms"
        Case MC_EMP_CCITT
            Value2String = "ccitt"
    End Select

Case MC_OUTPUT_FORMAT ' 26
    Select Case Param
        Case MC_OUTPUT_NORMAL
            Value2String = "mp3'"
        Case MC_OUTPUT_RIFF_WAVE
            Value2String = "riff/wave"
        Case MC_OUTPUT_RIFF_RMP
            Value2String = "riff/rmp"
    End Select
Case Else
    Value2String = "?"
End Select
End Function

Public Function InfoGogo(ByVal MG_param As Long) As Variant
    Dim value As Long
    
    If MG_param = MG_OUTPUTFILE Then
        Dim pun As Long, mp3File As String
        mp3File = Space$(50)
        pun = GlobalAlloc(GPTR, Len(mp3File))
        
        Resp = MPGE_getConfigureVB(MG_OUTPUTFILE, ByVal pun)
        GoSub verify
        CopyMemory mp3File, pun, Len(mp3File)
        InfoGogo = mp3File
        
    Else
    
        If MPGE_getConfigureVB(MG_param, value) = 0 Then
            InfoGogo = value
        Else
            InfoGogo = -1
        End If
        GoSub verify
    End If
    
    Exit Function
verify:
    If Resp <> 0 Then
        MsgBox GetError(CInt(Resp))
        GlobalFree pun
        Exit Function
    End If
    Return
 
End Function


Public Function Opciones_Frecuencia_GOGO(ByVal indice As Long) As Long
    indice = Abs(indice)
    Opciones_Frecuencia_GOGO = Choose((indice Mod 8) + 1, 44100, 32000, 24000, 22050, 16000, 12000, 11025, 8000)
End Function

Public Function Opciones_Kbps_GOGO(ByVal indice As Integer) As Long
    indice = Abs(indice)
    Opciones_Kbps_GOGO = Choose((indice Mod 14) + 1, 32, 40, 48, 56, 64, 80, 96, 112, 128, 160, 192, 224, 256, 320)
End Function

Public Function Opciones_Modo_GOGO(indice As Long, strResp As String) As Long
    indice = Abs(indice)
    strResp = CStr(Choose((indice Mod 5) + 1, "MONO", "STEREO", "JOINT", "msSTEREO", "DUAL"))
    Opciones_Modo_GOGO = indice Mod 5
End Function

Public Function Opciones_Enfasis_GOGO(indice As Long, strResp As String) As Long
    indice = Abs(indice)
    strResp = CStr(Choose((indice Mod 3) + 1, "No-EMPH", "50/15ms", "CCITT"))
    Opciones_Enfasis_GOGO = indice Mod 3
End Function

