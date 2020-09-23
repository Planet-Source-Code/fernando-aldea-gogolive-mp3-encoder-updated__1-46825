VERSION 5.00
Begin VB.Form frmOptions 
   Caption         =   "Options"
   ClientHeight    =   1335
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   3510
   LinkTopic       =   "Form3"
   ScaleHeight     =   1335
   ScaleWidth      =   3510
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1620
      TabIndex        =   4
      Top             =   660
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   660
      Width           =   915
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Text            =   "0"
      Top             =   180
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Bytes"
      Height          =   195
      Left            =   2880
      TabIndex        =   2
      Top             =   240
      Width           =   555
   End
   Begin VB.Label Label1 
      Caption         =   "Buffer Length:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''
''    Module written by Fernando Aldea G.   ''
''    e-mail: fernando_aldea@terra.cl       ''
''    Release January, 2003                 ''
''                                          ''
''    sorry for not translate completly     ''
''    & sorry for my English!               ''
''''''''''''''''''''''''''''''''''''''''''''''

Private Sub Command1_Click()
    BUFFER_LENGTH = Val(Me.Text1.Text)
    frmGogoLive.lblGOGO(1).Caption = BUFFER_LENGTH
    mPipe.Destruir
    If Not mPipe.Crear(BUFFER_LENGTH) Then MsgBox "Error en Pipe"
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    'frmGogoLive.Enabled = False
    Me.Text1.Text = BUFFER_LENGTH
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'frmGogoLive.Enabled = True
End Sub

Public Sub callBackWave(ByVal pBuffer As Long, largoBuffer As Long)
  On Error GoTo err
   
   Static tiempoAux As Long
   Beep
   Call mPipe.toWrite(pBuffer, largoBuffer)
   
   tiempoAux = DateDiff("s", TiempoInicio, Now)
   Me.Caption = Format((tiempoAux \ 60), "#00") & ":" & Format(tiempoAux Mod 60, "0#")
   
   Exit Sub
err:
   
End Sub


