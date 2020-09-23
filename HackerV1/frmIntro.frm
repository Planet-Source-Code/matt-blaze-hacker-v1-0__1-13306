VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Begin VB.Form frmIntro 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8610
   Icon            =   "frmIntro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   8610
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Left            =   7200
      Top             =   3000
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   255
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   30
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   0   'False
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   0
      WindowlessVideo =   0   'False
   End
   Begin VB.Timer Timer1 
      Left            =   7680
      Top             =   120
   End
   Begin VB.Label lblBomb3Info 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ACTIVE"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   7080
      TabIndex        =   10
      Top             =   1560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblBomb2Info 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ACTIVE"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2160
      TabIndex        =   9
      Top             =   600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblBomb1Info 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ACTIVE"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   3000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblBomb3 
      BackStyle       =   0  'Transparent
      Caption         =   "Bomb 3"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   7080
      TabIndex        =   7
      Top             =   1200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblBomb1 
      BackStyle       =   0  'Transparent
      Caption         =   "Bomb 1"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   2640
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblBomb2 
      BackStyle       =   0  'Transparent
      Caption         =   "Bomb 2"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblMission 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Mission Status:"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   2280
      TabIndex        =   2
      Top             =   960
      Width           =   4215
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblDisarm 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Disarm the three bombs..."
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   615
      Left            =   720
      TabIndex        =   0
      Top             =   1920
      Visible         =   0   'False
      Width           =   7095
   End
   Begin VB.Label lblMark 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   200.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   3735
      Left            =   3000
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   2415
   End
End
Attribute VB_Name = "frmIntro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Counter As Integer

Private Sub Form_Click()
Unload Me
frmL1Info.Show
End Sub

Private Sub Form_Load()
frmIntro.MediaPlayer1.filename = ("intro7.MP3")
Timer1.Interval = 1000
Counter = 0
End Sub

Private Sub imgContinue_Click()
Unload Me
frmMain.Show
End Sub

Private Sub Timer1_Timer()
Counter = Counter + 1
Label1 = Counter
If Counter = 3 Then
lblMission.Visible = False
frmIntro.Picture = LoadPicture("mov1.JPG")
End If
If Counter = 6 Then
frmIntro.Picture = LoadPicture("mov2.JPG")
End If
If Counter = 9 Then frmIntro.Picture = LoadPicture("mov3.JPG")
If Counter = 12 Then
lblBomb1.Visible = True
lblBomb2.Visible = True
lblBomb3.Visible = True
lblBomb1Info.Visible = True
lblBomb2Info.Visible = True
lblBomb3Info.Visible = True
End If
If Counter = 15 Then
lblBomb1.Visible = False
lblBomb2.Visible = False
lblBomb3.Visible = False
lblBomb1Info.Visible = False
lblBomb2Info.Visible = False
lblBomb3Info.Visible = False
frmIntro.Picture = LoadPicture("")
lblDisarm.Visible = True
End If

End Sub
