VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Begin VB.Form frmL1Info 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5670
   Icon            =   "frmL1Info.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   5670
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   6360
      Top             =   0
   End
   Begin VB.PictureBox P1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4260
      Left            =   480
      ScaleHeight     =   4200
      ScaleWidth      =   5595
      TabIndex        =   1
      Top             =   240
      Width           =   5655
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   135
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   135
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
      Filename        =   "briefing.MP3"
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
End
Attribute VB_Name = "frmL1Info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Dim thetop As Long
Dim p1hgt As Long
Dim p1wid As Long
Dim theleft As Long
Dim Tempstring As String

Private Sub Form_Click()
Unload Me
frmMain.Show
End Sub

Sub Form_Load()

        P1.AutoRedraw = True
        P1.Visible = False
        P1.FontSize = 10
        P1.ForeColor = &HFF00&
        P1.BackColor = BackColor
        P1.ScaleMode = 3
        ScaleMode = 3
        Open (App.Path & "\level1.txt") For Input As #1
        Line Input #1, Tempstring
        P1.Height = (Val(Tempstring) * P1.TextHeight("Test Height")) + 300
        Do Until EOF(1)
            Line Input #1, Tempstring
            PrintText Tempstring
        Loop
        Close #1
        theleft = 0
        thetop = ScaleHeight
        p1hgt = P1.ScaleHeight
        p1wid = P1.ScaleWidth
        Timer1.Enabled = True
        Timer1.Interval = 20
End Sub



Sub Timer1_Timer()
       x% = BitBlt(hDC, theleft, thetop, p1wid, p1hgt, P1.hDC, 0, 0, &HCC0020)
        thetop = thetop - 1
        If thetop < -p1hgt Then
        Timer1.Enabled = False
        Txt$ = "Credits Completed"
        CurrentY = ScaleHeight / 2
        CurrentX = (ScaleWidth - TextWidth(Txt$)) / 2
        Print Txt$
        End If
End Sub

Sub PrintText(Text As String)
P1.CurrentX = (P1.ScaleWidth / 2) - (P1.TextWidth(Text) / 2)
P1.ForeColor = 0: x = P1.CurrentX: y = P1.CurrentY
For i = 1 To 3
    P1.Print Text
    x = x + 1: y = y + 1: P1.CurrentX = x: P1.CurrentY = y
Next i
P1.ForeColor = &HFF00&
P1.Print Text
End Sub


