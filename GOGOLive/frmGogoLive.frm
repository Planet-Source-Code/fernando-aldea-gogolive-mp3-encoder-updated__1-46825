VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGogoLive 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GOGOLive-Fernando Aldea Jul/2003"
   ClientHeight    =   6240
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   3735
   ForeColor       =   &H8000000D&
   Icon            =   "frmGogoLive.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   3735
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2160
      Top             =   4920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   13
      ImageHeight     =   13
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGogoLive.frx":0442
            Key             =   "Rec"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGogoLive.frx":069E
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGogoLive.frx":08FA
            Key             =   "Pause"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGogoLive.frx":0B56
            Key             =   "CD"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGogoLive.frx":0E12
            Key             =   "Wave"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGogoLive.frx":10CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGogoLive.frx":132A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGogoLive.frx":1586
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGogoLive.frx":17E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGogoLive.frx":1A3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGogoLive.frx":1C9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGogoLive.frx":1EF6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sbInfo 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   26
      Top             =   5985
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Display 
      Align           =   1  'Align Top
      BackColor       =   &H00000000&
      Height          =   1635
      Left            =   0
      ScaleHeight     =   1575
      ScaleWidth      =   3675
      TabIndex        =   1
      Top             =   375
      Width           =   3735
      Begin VB.PictureBox Spectrum 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         DragMode        =   1  'Automatic
         FillStyle       =   0  'Solid
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   900
         Negotiate       =   -1  'True
         ScaleHeight     =   285
         ScaleWidth      =   1605
         TabIndex        =   11
         Top             =   660
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.Label lblDer 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "<<<<>>>>"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   0
         Left            =   2640
         TabIndex        =   25
         Tag             =   "0"
         ToolTipText     =   "Click here to change Frequency"
         Top             =   300
         Width           =   975
      End
      Begin VB.Label lblDer 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "<><><>"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   1
         Left            =   2640
         TabIndex        =   24
         Tag             =   "0"
         ToolTipText     =   "Click here to change Bitrate"
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lblDer 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "<><><>"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   2
         Left            =   2640
         TabIndex        =   23
         Tag             =   "0"
         ToolTipText     =   "Click here to change Channels mode"
         Top             =   900
         Width           =   975
      End
      Begin VB.Label lblIzq 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PSY"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   0
         Left            =   60
         TabIndex        =   22
         ToolTipText     =   "Psycho-acustic mode. Best Quality"
         Top             =   300
         Width           =   735
      End
      Begin VB.Label lblIzq 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LPF16"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   1
         Left            =   60
         TabIndex        =   21
         ToolTipText     =   "Low Pass Filter mode."
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblIzq 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MMX"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   2
         Left            =   60
         TabIndex        =   20
         ToolTipText     =   "Use MMX capacity if is available"
         Top             =   900
         Width           =   735
      End
      Begin VB.Label LCD1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "READY"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Left            =   900
         TabIndex        =   19
         ToolTipText     =   "Current state"
         Top             =   300
         Width           =   1695
      End
      Begin VB.Label LCD2 
         Alignment       =   2  'Center
         BackColor       =   &H00400000&
         Caption         =   "00:00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   315
         Left            =   900
         TabIndex        =   18
         Top             =   660
         Width           =   1695
      End
      Begin VB.Label lblSong 
         BackStyle       =   0  'Transparent
         Caption         =   "- - - - - - -"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Width           =   3735
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblGOGO 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "GOGO"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   0
         Left            =   840
         TabIndex        =   16
         ToolTipText     =   "Gogo.dll version"
         Top             =   1020
         Width           =   1755
      End
      Begin VB.Label lblDer 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "<><><>"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   3
         Left            =   2640
         TabIndex        =   15
         Tag             =   "0"
         ToolTipText     =   "Click here to change Emphasys mode"
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label lblIzq 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CPU"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   3
         Left            =   60
         TabIndex        =   14
         ToolTipText     =   "Use multiple CPU if is available"
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label lblGOGO 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "GOGOLive"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   195
         Index           =   1
         Left            =   840
         TabIndex        =   13
         ToolTipText     =   "Gogo.dll version"
         Top             =   1260
         Width           =   1755
      End
      Begin VB.Label LCD3 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         Caption         =   "!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   2340
         TabIndex        =   12
         ToolTipText     =   "Low processing PC"
         Top             =   420
         Visible         =   0   'False
         Width           =   255
      End
   End
   Begin VB.PictureBox DisplayCD 
      Align           =   1  'Align Top
      BackColor       =   &H00000000&
      Height          =   1335
      Left            =   0
      ScaleHeight     =   1275
      ScaleWidth      =   3675
      TabIndex        =   2
      Top             =   2385
      Width           =   3735
      Begin VB.ListBox ListTracks 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         ForeColor       =   &H00FFFFFF&
         Height          =   930
         Left            =   1800
         Style           =   1  'Checkbox
         TabIndex        =   5
         ToolTipText     =   "List CD tracks"
         Top             =   300
         Width           =   1815
      End
      Begin VB.Label lblSongCD 
         BackStyle       =   0  'Transparent
         Caption         =   "- - - - - - -"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   10
         ToolTipText     =   "name track playing"
         Top             =   0
         Width           =   3735
      End
      Begin VB.Label LCD2_CD 
         Alignment       =   2  'Center
         BackColor       =   &H00400000&
         Caption         =   "00:00:00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   315
         Left            =   60
         TabIndex        =   9
         ToolTipText     =   "Time of track"
         Top             =   660
         Width           =   1755
      End
      Begin VB.Label LCD1_CD 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "READY"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Left            =   60
         TabIndex        =   8
         ToolTipText     =   "Current CD state "
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label lbl2Izq 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "TIMER"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   60
         TabIndex        =   7
         Tag             =   "1"
         ToolTipText     =   "Show Time of track ON/OFF"
         Top             =   1020
         Width           =   615
      End
      Begin VB.Label miniLCD_CD 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Caption         =   "00%"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1200
         TabIndex        =   6
         ToolTipText     =   "Relative current position of track"
         Top             =   360
         Width           =   615
      End
   End
   Begin MSComctlLib.Toolbar tbControlCD 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   2010
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   661
      ButtonWidth     =   529
      ButtonHeight    =   503
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Refresh"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Play"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Stop"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Prev"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Next"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Rec-Play"
            ImageIndex      =   12
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbControlMP3 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   661
      ButtonWidth     =   529
      ButtonHeight    =   503
      Wrappable       =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "New mp3 file"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Rec"
            Object.ToolTipText     =   "Start record to mp3"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Stop"
            Object.ToolTipText     =   "Stop current recording"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Pause"
            Object.ToolTipText     =   "Pause currect recording"
            ImageIndex      =   3
            Style           =   1
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Monitor"
            Object.ToolTipText     =   "Monitor input level"
            ImageIndex      =   6
            Style           =   1
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   1440
      Picture         =   "frmGogoLive.frx":2186
      ScaleHeight     =   135
      ScaleWidth      =   1470
      TabIndex        =   3
      Top             =   4320
      Width           =   1530
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   1020
      Top             =   4680
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   480
      Top             =   4680
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   0
      Top             =   4680
   End
   Begin VB.Menu mnArchivo 
      Caption         =   "File"
      Begin VB.Menu mnNuevo 
         Caption         =   "New..."
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnVer 
      Caption         =   "&View"
      Begin VB.Menu mnuControlCD 
         Caption         =   "Show Control CD"
      End
      Begin VB.Menu mnuOpciones 
         Caption         =   "&Options"
      End
   End
   Begin VB.Menu About 
      Caption         =   "About"
   End
End
Attribute VB_Name = "frmGogoLive"
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

Const AMPLITUD_MAX = 65535

Private Sub About_Click()
    frmAbout.Show
End Sub




Sub CmdMonitor()
    If Me.tbControlMP3.Buttons("Monitor").value = 1 Then
        
        If Not Monitoring Then
            If mWaveCard.iniMonitor() = True Then
                frmGogoLive.LCD1.Caption = "MONITOR"
                frmGogoLive.LCD2.Visible = False
                frmGogoLive.Spectrum.Visible = True
                Me.tbControlMP3.Buttons("Rec").Enabled = False
                Me.tbControlMP3.Buttons("Stop").Enabled = False
                Me.tbControlMP3.Buttons("Pause").Enabled = False
                Monitoring = True
            Else
                Me.tbControlMP3.Buttons("Monitor").value = 0
                frmGogoLive.LCD1.Caption = "Error"
                Monitoring = False
                
            End If
        End If
     Else
        mWaveCard.TerminarMonitor
        Me.tbControlMP3.Buttons("Rec").Enabled = True
        Me.tbControlMP3.Buttons("Stop").Enabled = True
        Me.tbControlMP3.Buttons("Pause").Enabled = True
        frmGogoLive.LCD2.Visible = True
        frmGogoLive.Spectrum.Visible = False
        Me.LCD1.Caption = "READY"
        Monitoring = False
    End If
    
End Sub


Private Sub CmdPause()
    
    If Me.tbControlMP3.Buttons("Pause").value = 1 And mWaveCard.isPausing Then Exit Sub
    If Me.tbControlMP3.Buttons("Pause").value = 0 And Not mWaveCard.isPausing Then Exit Sub
    
    PauseRecord
    
    If mWaveCard.isPausing Then
        Me.LCD1.Caption = "PAUSE"
       ' Me.CmdPause.Value = 1
    Else
        'Me.CmdPause.Value = 0
        If Recording Then
            Me.LCD1.Caption = "REC"
        Else
            Me.LCD1.Caption = "READY"
        End If
    End If
    
End Sub

Sub CmdPlayCD()
    Me.LCD1_CD.Caption = "PLAY..."
    Me.LCD1_CD.Refresh
    mCDaudio.PlayCD
    Me.LCD1_CD.Caption = "PLAY"
    If Me.lbl2Izq.Tag = 1 Then Me.Timer2.Enabled = True
End Sub

Private Sub CmdRec()

    If Not Recording Then
        
        If Not mPipe.estaListo Then MsgBox "error Pipe": Exit Sub
        
        Me.tbControlMP3.Buttons("Rec").Enabled = False
        Me.tbControlMP3.Buttons("New").Enabled = False
        Me.tbControlMP3.Buttons("Stop").Enabled = True
        Me.tbControlMP3.Buttons("Monitor").Enabled = False
        Me.lblDer(0).Enabled = False
        Me.lblDer(1).Enabled = False
        Me.lblDer(2).Enabled = False
        Me.lblDer(3).Enabled = False
        Me.lblIzq(0).Visible = False
        Me.lblIzq(1).Visible = False
        Me.lblIzq(2).Visible = False
        Me.lblIzq(3).Visible = False
        Me.lblIzq(0).Enabled = False
        Me.lblIzq(1).Enabled = False
        Me.lblIzq(2).Enabled = False
        Me.lblIzq(3).Enabled = False
        Me.Timer1.Enabled = True
        
        Me.LCD1.ForeColor = &HFF&
        Me.LCD1.Caption = "REC"
        'Me.lblGOGO(1).Caption = "Buffer:" & BUFFER_LENGTH
        nBytesPCMacum = 0
        
        StartRecord
        
        If Recording Then
            Me.LCD1.Caption = "ERROR"
            Me.LCD1.ForeColor = &HFFFFFF
        Else
            Me.LCD1.Caption = "READY"
            Me.LCD1.ForeColor = &HFFFFFF
        End If
        
        EndRecord
        Timer1_Timer
        Me.Timer1.Enabled = False
        Me.tbControlMP3.Buttons("Rec").Enabled = True
        Me.tbControlMP3.Buttons("New").Enabled = True
        'Me.tbControlMP3.Buttons("Stop").Enabled = False
        Me.tbControlMP3.Buttons("Monitor").Enabled = True
        Me.lblDer(0).Enabled = True
        Me.lblDer(1).Enabled = True
        Me.lblDer(2).Enabled = True
        Me.lblDer(3).Enabled = True
        Me.lblIzq(0).Visible = True
        Me.lblIzq(1).Visible = True
        Me.lblIzq(2).Visible = True
        Me.lblIzq(3).Visible = True
        Me.lblIzq(0).Enabled = True
        Me.lblIzq(1).Enabled = True
        Me.lblIzq(2).Enabled = True
        Me.lblIzq(3).Enabled = True
               
        
    End If
    
End Sub


Sub CmdRecPlay()
    Me.tbControlCD.Buttons("Rec-Play").Enabled = False
    If Me.ListTracks.ListCount = 0 Then
        MsgBox "Nada seleccionado"
    Else
        'gravar la lista
        For nt = 0 To Me.ListTracks.ListCount - 1
            If Me.ListTracks.Selected(nt) Then
                Me.ListTracks.ListIndex = nt
                Call ListTracks_MouseDown(2, 0, 0, 0)
                
                MP3_outFile = mCDaudio.lstTracks(nt + 1).Nombre
                Me.lblSong = MP3_outFile
                
                Me.tbControlMP3.Buttons("Pause").value = 1
                Me.Timer3.Enabled = True
                Call CmdRec
                Me.Timer3.Enabled = False
                
                Me.ListTracks.List(nt) = mCDaudio.lstTracks(nt + 1).Nombre & " (OK)"
                Me.ListTracks.Selected(nt) = False
            End If
        Next nt
    End If
    Me.tbControlCD.Buttons("Rec-Play").Enabled = True
End Sub

Sub CmdRefreshCD()
     
    Me.ListTracks.Clear
    Me.LCD1_CD.Caption = "WAIT..."
    Me.LCD1_CD.Refresh
    If mCDaudio.Actualizar() Then
        For a = 1 To mCDaudio.numTracks
            Me.ListTracks.AddItem mCDaudio.lstTracks(a).Nombre & " (" & mCDaudio.lstTracks(a).Largo & ")"
        Next a
        Me.LCD1_CD.Caption = "READY"
    Else
        Me.LCD1_CD.Caption = "ERROR"
    End If
    
    
End Sub

Private Sub CmdStop()
     If Recording Then
        Me.LCD1.Caption = "STOP..."
        EndRecord
        Me.tbControlMP3.Buttons("Pause").value = 0
    End If
End Sub






Sub CmdStopCD()
    Me.LCD1_CD.Caption = "STOP..."
    Me.LCD1_CD.Refresh
    mCDaudio.StopCD
    Me.LCD1_CD.Caption = "STOP"
    Me.Timer2.Enabled = False
End Sub



Private Sub Form_Load()
    Dim mStr As String
    
    
    
    Me.lblSong.Caption = MP3_outFile
    Me.lblGOGO(0).Caption = Version()
    Me.lblGOGO(1).Caption = "GOGOLive " & App.Major & "." & App.Minor
    
    Me.lblDer(0).Caption = MP3_Frequency & " Hz"
    Me.lblDer(1).Caption = MP3_Kbps & " Kbps"
            
    Call Opciones_Modo_GOGO(MP3_Mode, mStr)
    Me.lblDer(2).Caption = mStr
    
    Call Opciones_Enfasis_GOGO(MP3_Enphasis, mStr)
    Me.lblDer(3).Caption = mStr
    
    
        If MP3_PSY Then Me.lblIzq(0).Caption = "PSY" Else Me.lblIzq(0).Caption = ""
        If MP3_LPF16 Then Me.lblIzq(1).Caption = "LPF16" Else Me.lblIzq(1).Caption = ""
        If MP3_MMX Then Me.lblIzq(2).Caption = "MMX" Else Me.lblIzq(2).Caption = ""
        If MP3_CPU Then Me.lblIzq(3).Caption = "CPU" Else Me.lblIzq(3).Caption = ""
    
    If Not mPipe.Crear(BUFFER_LENGTH) Then MsgBox "No se pudo crear Pipe", vbCritical
    
    ModoNoCD
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    If Recording Then EndRecord
    mPipe.Destruir
    mCDaudio.FinalizarCD
End Sub




Private Sub lbl2Izq_Click()
    If Me.lbl2Izq.Tag = 0 Then
        Me.lbl2Izq.Caption = "TIMER"
        Me.lbl2Izq.Tag = 1
        Me.Timer2.Enabled = True
        Me.LCD2_CD.Enabled = True
    Else
        Me.lbl2Izq.Caption = ""
        Me.lbl2Izq.Tag = 0
        Me.Timer2.Enabled = False
        Me.LCD2_CD.Caption = "NONE"
        Me.LCD2_CD.Enabled = False
    End If
End Sub

Private Sub lblDer_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim mStr As String
    If Button = 1 Then
        lblDer(Index).Tag = lblDer(Index).Tag + 1
    Else
        lblDer(Index).Tag = lblDer(Index).Tag - 1
    End If
    
    If Index = 0 Then
        MP3_Frequency = Opciones_Frecuencia_GOGO(lblDer(Index).Tag)
        Me.lblDer(Index).Caption = MP3_Frequency & " Hz"
    End If
    
    If Index = 1 Then
        MP3_Kbps = Opciones_Kbps_GOGO(lblDer(Index).Tag)
        Me.lblDer(Index).Caption = MP3_Kbps & " Kbps"
    End If
    
    If Index = 2 Then
        MP3_Mode = Opciones_Modo_GOGO(Me.lblDer(Index).Tag, mStr)
        Me.lblDer(Index).Caption = mStr
    End If
        
    If Index = 3 Then
        MP3_Enphasis = Opciones_Enfasis_GOGO(lblDer(Index).Tag, mStr)
        Me.lblDer(Index).Caption = mStr
    End If
    
End Sub

Private Sub lblIzq_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Index = 0 Then
        MP3_PSY = Not MP3_PSY
        If MP3_PSY Then Me.lblIzq(Index).Caption = "PSY" Else Me.lblIzq(Index).Caption = ""
    End If
    If Index = 1 Then
        MP3_LPF16 = Not MP3_LPF16
        If MP3_LPF16 Then Me.lblIzq(Index).Caption = "LPF16" Else Me.lblIzq(Index).Caption = ""
    End If
    If Index = 2 Then
        MP3_MMX = Not MP3_MMX
        If MP3_MMX Then Me.lblIzq(Index).Caption = "MMX" Else Me.lblIzq(Index).Caption = ""
    End If
    If Index = 3 Then
        MP3_CPU = Not MP3_CPU
        If MP3_CPU Then Me.lblIzq(Index).Caption = "CPU" Else Me.lblIzq(Index).Caption = ""
    End If
End Sub


Private Sub lblSong_Click()
    Dim mStr As String
    
    mStr = InputBox("Nuevo nombre de archivo", , MP3_outFile)
    If mStr <> "" Then
        MP3_outFile = mStr
    End If
    Me.lblSong = MP3_outFile
    
End Sub


Private Sub ListTracks_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then Exit Sub
    If Me.ListTracks.ListCount = 0 Then Exit Sub
    Me.LCD1_CD.Caption = "SEEK..."
    Me.LCD1_CD.Refresh
    
    If mCDaudio.SetTrackActual(Me.ListTracks.ListIndex + 1) Then
        Me.lblSongCD.Caption = mCDaudio.numTrackActual & ".-" & mCDaudio.lstTracks(mCDaudio.numTrackActual).Nombre & _
                                " (" & mCDaudio.lstTracks(mCDaudio.numTrackActual).Largo & ")"
        Call Timer2_Timer
        Me.LCD1_CD.Caption = "READY"
    Else
        Me.LCD1_CD.Caption = "ERROR"
        MsgBox UltimoError(), vbCritical
    End If
    
End Sub

Private Sub mnNuevo_Click()
    cmdNuevo
End Sub

Private Sub mnuControlCD_Click()
    mnuControlCD.Checked = Not mnuControlCD.Checked
    If mnuControlCD.Checked Then
        ModoCD
    Else
        ModoNoCD
    End If
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuOpciones_Click()
    frmOptions.Show
End Sub

Sub cmdNuevo()
    Dim Resp As String
        
    Resp = InputBox("Enter new output mp3 file:", "New Mp3", OUTFILE_DEFAULT)
    If Resp = "" Then Exit Sub
    MP3_outFile = Resp
    Me.lblSong.Caption = MP3_outFile
    Me.LCD2.Caption = "00:00"
    Me.sbInfo.Panels(1).Text = ""
    Me.sbInfo.Panels(2).Text = ""
End Sub

Public Sub callBackWave(ByVal pBuffer As Long, largoBuffer As Long)
  On Error GoTo err
  
  If Monitoring And Not Recording Then
    Monitor pBuffer, largoBuffer
    Exit Sub
  End If
  
  If nBytesPCMacum = -1 Then
    TimerIni = Timer
    nBytesPCMacum = 0
  Else
    nBytesPCMacum = nBytesPCMacum + largoBuffer
  End If

  
  If Abs((Timer - TimerIni) - (nBytesPCMacum / mWaveCard.nAvgBytesPerSec)) >= 0.1 Then
        Me.LCD3.Visible = True
  Else
        Me.LCD3.Visible = False
  End If
        
   Call mPipe.toWrite(pBuffer, largoBuffer)
  
  
   Exit Sub
err:
    'Me.LCD1.Caption = "Error"
End Sub

Public Function callBackMp3(ByVal pBuffer, largoBuffer As Long) As Long
  On Error GoTo err
   
   Dim LeidosPipe As Long
   Static BytesTotal As Long
   Static BytesActual As Long
   
   BytesTotal = mPipe.nBytesTotal
   BytesActual = mPipe.nBytesActual
      
    
   'If mPipe.nBytesTotal = 0 Then 'pipe infinito
   '      Me.sbInfo.Panels(2).text = Format$(BytesActual / 1024, "###.0") & " KB"
   'Else
   '      Me.sbInfo.Panels(2).text = Format$((BytesTotal - BytesActual) * 100 / BytesTotal, "###.0") & "%"
   'End If
   
   'Me.sbInfo.Panels(1).text = curFrame & " frames"
      
   'condicion de salida
   If Not Recording And BytesActual <= 0 Then callBackMp3 = ME_EMPTYSTREAM: Exit Function
   
   'esperar si el pipe no esta lleno lo suficiente
   If Recording And (BytesActual < largoBuffer) Then callBackMp3 = ME_MOREDATA: Exit Function
    
    'toRead pipe
    LeidosPipe = mPipe.toRead(pBuffer, largoBuffer)
        
    If LeidosPipe < largoBuffer And Not Recording Then
        callBackMp3 = ME_EMPTYSTREAM
    Else
        callBackMp3 = ME_NOERR
    End If
     
    
   Exit Function
err:
   callBackMp3 = ME_INTERNALERROR
   'Me.sbInfo.Panels(1).Text = "Error GOGO"
      
End Function

Private Sub tbControlCD_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Refresh"
            CmdRefreshCD
        Case "Play"
            CmdPlayCD
        Case "Stop"
            CmdStopCD
        Case "Prev"
            'cmdPrevCD
        Case "Next"
            'cmdNextCD
        Case "Rec-Play"
            CmdRecPlay
    End Select
    
    
End Sub

Private Sub Timer1_Timer()

    Static tiempoAux As Long

    Me.sbInfo.Panels(1).Text = "Mp3: " & curFrame & " frames"
    
    If mPipe.nBytesTotal = 0 Then 'pipe infinito
         Me.sbInfo.Panels(2).Text = Format$(mPipe.nBytesActual / 1024, "##0.0") & " KB"
   Else
         Me.sbInfo.Panels(2).Text = "Buffer: " & Format$((mPipe.nBytesTotal - mPipe.nBytesActual) * 100 / mPipe.nBytesTotal, "##0.0") & "% free"
   End If
   
  
  
   tiempoAux = CLng(nBytesPCMacum \ mWaveCard.nAvgBytesPerSec)
   
   Me.LCD2.Caption = Format$((tiempoAux \ 60), "#00") & ":" & Format(tiempoAux Mod 60, "0#")
  
       
End Sub

Private Sub Timer2_Timer()
    Me.LCD2_CD.Caption = mCDaudio.Position()
End Sub

Sub ModoNoCD()
    mCDaudio.FinalizarCD
        
    Me.DisplayCD.Visible = False
    Me.tbControlCD.Visible = False
    Me.Height = 2925
End Sub
Sub ModoCD()
    If Not mCDaudio.InitializeCD() Then MsgBox "No CD present!", vbCritical: Exit Sub
    
    Me.tbControlCD.Visible = True
    Me.DisplayCD.Visible = True
    Me.Height = 4650
End Sub

Private Sub Timer3_Timer()
    
    Static tpoInicio As Single
    'Static tpoMeta As Long
    Static Dif As Long
    
    
    If Recording And mWaveCard.isPausing And mCDaudio.CDEstaListo Then
        Me.tbControlMP3.Buttons("Pause").value = 0
        Me.Timer3.Interval = 500
        CmdPlayCD
        tpoInicio = Timer
        Me.Timer3.Enabled = True
    End If
    If Not Recording Then
        Me.Timer3.Enabled = False
    End If
    
    Dif = CLng((Timer - tpoInicio) * 1000)
    If Dif >= mCDaudio.lstTracks(numTrackActual).Miliseg Then
        Me.Timer3.Enabled = False
        CmdStop
        CmdStopCD
    End If
    Me.miniLCD_CD.Caption = Format$(Dif * 100 \ mCDaudio.lstTracks(numTrackActual).Miliseg, "##0") & "%"
End Sub



Sub Monitor(ByVal pWave As Long, Largo As Long)
On Error GoTo err
    Static waveData(1 To 4) As Byte
    Static L As Long
    Static R As Long
    Static aL As Long
    Static aR As Long
    
    CopyMemory waveData(1), ByVal pWave, 4 'Largo
    
    If wformat.nChannels = 2 Then
        If wformat.wBitsPerSample = 16 Then
            L = Amplitud(waveData(1), waveData(2))
            R = Amplitud(waveData(3), waveData(4))
        Else
            L = Amplitud(waveData(1))
            R = Amplitud(waveData(2))
        End If
    Else
        If wformat.wBitsPerSample = 16 Then
            L = Amplitud(waveData(1), waveData(2))
        Else
            L = Amplitud(waveData(1))
        End If
        
    End If
    
    L = Abs(L)
    R = Abs(R)
    
    If aL - L > 500 Then L = aL - 500
    If aR - R > 500 Then R = aR - 500
        
    'Me.Shape1.Width = Abs(L) * 1515 / 33000
    'Me.Shape2.Width = Abs(R) * 1515 / 33000
    
    
   ' BitBlt Me.Spectrum.hDC, 0, 0, Me.Width, Me.Height, Me.Picture2.hDC, (Abs(R) * Me.Picture2.Width / 33000), (Abs(R) * Me.Picture2.Height / 33000), SRCCOPY
    'BitBlt Me.Picture3, 0, 0, Me.Picture3.Width, Me.Picture3.Height, Me.Picture2.hDC, 0, 0, SRCCOPY
    Me.Spectrum.Cls
    BitBlt Me.Spectrum.hDC, 0, 0, L * Me.Picture2.ScaleX(Me.Picture2.Width) / 65000, Me.Picture2.ScaleY(Me.Picture2.Height), Me.Picture2.hDC, 0, 0, SRCCOPY
    BitBlt Me.Spectrum.hDC, 0, 10, R * Me.Picture2.ScaleX(Me.Picture2.Width) / 65000, Me.Picture2.ScaleY(Me.Picture2.Height), Me.Picture2.hDC, 0, 0, SRCCOPY
    
    aL = L
    aR = R
    
    'Me.Picture2.Refresh
    Exit Sub
err:

End Sub

Function Amplitud(Byte1 As Byte, Optional Byte2 As Byte) As Long
    'Dim Res As Long
    
    Amplitud = Byte1 + (Byte2 * (2 ^ 8))

    'If Byte2 = 0 Then Exit Sub
    
    If Byte2 >= 128 Then
        Amplitud = Amplitud - 65536
    End If
    
End Function


Private Sub tbControlMP3_ButtonClick(ByVal Button As MSComctlLib.Button)
     Select Case Button.Key
        Case "New"
            cmdNuevo
        Case "Rec"
            CmdRec
        Case "Stop"
            CmdStop
        Case "Pause"
            CmdPause
        Case "Monitor"
            CmdMonitor
    End Select
End Sub
