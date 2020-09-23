VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Begin VB.Form frmWin 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8610
   Icon            =   "frmWin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   8610
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Left            =   7680
      Top             =   120
   End
   Begin VB.Timer Timer2 
      Left            =   7200
      Top             =   3000
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   255
      Left            =   120
      TabIndex        =   0
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
   Begin VB.Label lblRestart 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Restart"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2640
      MouseIcon       =   "frmWin.frx":000C
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   3840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblExit 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4680
      MouseIcon       =   "frmWin.frx":015E
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   3840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblDisarm 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Mission Successful"
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
      TabIndex        =   9
      Top             =   1920
      Visible         =   0   'False
      Width           =   7095
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
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
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
      TabIndex        =   7
      Top             =   960
      Width           =   4215
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
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   240
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
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   2640
      Visible         =   0   'False
      Width           =   1335
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
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   7080
      TabIndex        =   4
      Top             =   1200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblBomb1Info 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Disarmed"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   3000
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lblBomb2Info 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Disarmed"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lblBomb3Info 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Disarmed"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   6840
      TabIndex        =   1
      Top             =   1560
      Visible         =   0   'False
      Width           =   1695
   End
End
Attribute VB_Name = "frmWin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Flash, Counter As Integer


Private Sub Form_Load()
frmWin.MediaPlayer1.filename = "intro7.MP3"
Timer1.Enabled = True
Timer1.Interval = 1000
Counter = 0
Timer2.Enabled = False
Timer2.Interval = 1000
Flash = 0
End Sub


Private Sub lblExit_Click()
Unload Me
End
End Sub

Private Sub lblRestart_Click()
Unload Me
frmL1Info.Show
End Sub

Private Sub Timer1_Timer()
Counter = Counter + 1
Label1 = Counter
If Counter = 3 Then
lblMission.Visible = False
frmWin.Picture = LoadPicture("mov1.JPG")
End If
If Counter = 6 Then
frmWin.Picture = LoadPicture("mov2.JPG")
End If
If Counter = 9 Then frmWin.Picture = LoadPicture("mov3.JPG")
If Counter = 12 Then
Timer2.Enabled = True
lblBomb1.Visible = True
lblBomb2.Visible = True
lblBomb3.Visible = True
lblBomb1Info.Visible = True
lblBomb2Info.Visible = True
lblBomb3Info.Visible = True
End If
If Counter = 17 Then
lblBomb1.Visible = False
lblBomb2.Visible = False
lblBomb3.Visible = False
lblBomb1Info.Visible = False
lblBomb2Info.Visible = False
lblBomb3Info.Visible = False
frmWin.Picture = LoadPicture("")
lblDisarm.Visible = True
End If
If Counter = 19 Then
lblRestart.Visible = True
lblExit.Visible = True
End If
End Sub

Private Sub Timer2_Timer()
Flash = Flash + 1
If Flash = 1 Then
lblBomb1Info.Visible = False
lblBomb2Info.Visible = False
lblBomb3Info.Visible = False
End If
If Flash = 2 Then
lblBomb1Info.Visible = True
lblBomb2Info.Visible = True
lblBomb3Info.Visible = True
End If
If Flash = 3 Then
lblBomb1Info.Visible = False
lblBomb2Info.Visible = False
lblBomb3Info.Visible = False
End If
If Flash = 4 Then
lblBomb1Info.Visible = True
lblBomb2Info.Visible = True
lblBomb3Info.Visible = True
End If
End Sub
