Attribute VB_Name = "mCDaudio"
''''''''''''''''''''''''''''''''''''''''''''''
''    Module adapted by Fernando Aldea G.   ''
''    e-mail: fernando_aldea@terra.cl       ''
''    web: orbita.starmedia.com/gogolive/   ''
''    Release Juny, 2003                    ''
''                                          ''
''    sorry for not translate this completly''
''    & sorry about my English!             ''
''''''''''''''''''''''''''''''''''''''''''''''



Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long

Type InfoTrack
    n As Integer
    Nombre As String
    Largo As String
    Miliseg As Long
End Type

Public CDId As String
Public CDEstaListo As Boolean
Public numTracks As Integer
Public numTrackActual As Integer
Public lstTracks() As InfoTrack
Dim lRet As Long

Private Function CString(aStr As String) As String
    CString = ""
    Dim k As Long
    k = InStr(aStr, Chr$(0))
    If k Then
        CString = Left$(aStr, k - 1)
    End If
End Function


Public Function InitializeCD()
    If mciSendString("open cdaudio alias cd", vbNullString, 0, 0) = 0 Then
        CDEstaListo = True
        'Actualizar
    Else
        CDEstaListo = False
    End If
    InitializeCD = CDEstaListo
    numTrackActual = 1
End Function

Public Function FinalizarCD()
    mciSendString "close cd", vbNullString, 0, 0
    CDEstaListo = False
    FinalizarCD = CDEstaListo
    numTrackActual = 0
End Function



Public Function Actualizar() As Boolean
    Dim aRet As String, aTrack As String
    'Dim mTracks As InfoTrack
    
    aRet = Space$(64)
    aTrack = Space$(2)
    
    If CDEstaListo Then
        mciSendString "info cd identity", aRet, 64, 0
        CDId = CString(aRet)
        'txtFile.Text = App.Path & "\CD-" & lblCDID.Caption
        mciSendString "status cd number of tracks", aRet, 64, 0
        numTracks = Val(aRet)
        If numTracks = 0 Then Actualizar = True: Exit Function
        ReDim lstTracks(1 To numTracks)
        
        
        mciSendString "set cd time format hms", vbNullString, 0, 0
        For lRet = 1 To numTracks
            mciSendString "status cd length track " & lRet, aRet, 64, 0
            RSet aTrack = CStr(lRet)
            lstTracks(lRet).Nombre = "Track " & aTrack
            lstTracks(lRet).Largo = CString(aRet)
            lstTracks(lRet).n = lRet
        Next
        'ReDim lTrackLengths(1 To Val(lblNumTracks.Caption)) As Long
        mciSendString "set cd time format milliseconds", vbNullString, 0, 0
        For lRet = 1 To numTracks
            mciSendString "status cd length track " & lRet, aRet, 64, 0
            lstTracks(lRet).Miliseg = CLng(CString(aRet))
        Next
        
        Actualizar = True
    End If
End Function


Public Function PlayCD(Optional nTrack As Integer) As Boolean
    'mciSendString "set cd time format hms", vbNullString, 0, 0
    PlayCD = False
    If nTrack = 0 Or nTrack = numTrackActual Then
        lRet = mciSendString("play cd", vbNullString, 0, 0): Exit Function
        If lRet = 0 Then PlayCD = True
    Else
        If nTrack < numTracks Then
            lRet = mciSendString("play cd from " & nTrack & " to " & nTrack + 1, vbNullString, 0, 0)
            If lRet = 0 Then PlayCD = True
        Else
            lRet = mciSendString("play cd from " & nTrack, vbNullString, 0, 0)
            If lRet = 0 Then PlayCD = True
        End If
    End If
    
End Function

Public Function StopCD() As Boolean
    mciSendString "stop cd", vbNullString, 0, 0
End Function

Public Function EjectCD() As Boolean

End Function

Public Function Position() As String
Dim aRet As String ', lTrack As Long
    aRet = Space$(64)
    mciSendString "status cd position", aRet, 64, 0
    'lRet = Val(CString(aRet))
    Position = aRet
End Function

Public Function SetTrackActual(nTrack As Integer) As Boolean
    SetTrackActual = False
    If nTrack <= numTracks Then
        lRet = mciSendString("set cd time format tmsf", vbNullString, 0, 0)
        lRet = mciSendString("seek cd to " & CStr(nTrack), vbNullString, 0, 0)
        'lRet = mciSendString("set cd time format hms", vbNullString, 0, 0)
        If lRet = 0 Then
            numTrackActual = nTrack
            SetTrackActual = True
        End If
    End If
    
End Function

Public Function UltimoError() As String
    Dim mStr As String
    mStr = Space$(100)
    
    mciGetErrorString lRet, mStr, Len(mStr)
    UltimoError = mStr
End Function
