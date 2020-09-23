VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Terminal Console"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8055
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   8055
   StartUpPosition =   2  'CenterScreen
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   135
      Left            =   480
      TabIndex        =   63
      Top             =   1680
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
      Filename        =   "mainlong.mp3"
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   0
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
   Begin VB.Timer TimerWin 
      Enabled         =   0   'False
      Left            =   7320
      Top             =   3360
   End
   Begin VB.Timer TimerLose 
      Left            =   7320
      Top             =   2520
   End
   Begin VB.Timer TimerSC 
      Left            =   7560
      Top             =   120
   End
   Begin VB.Timer TimerX 
      Left            =   1560
      Top             =   4440
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   58
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Timer Timer4 
      Left            =   1560
      Top             =   0
   End
   Begin VB.Timer Timer3 
      Left            =   120
      Top             =   0
   End
   Begin VB.Timer Timer2 
      Left            =   600
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Left            =   1080
      Top             =   0
   End
   Begin VB.Label lblActivate2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "!! ACTIVATED !!"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   2160
      TabIndex        =   62
      Top             =   2520
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Label lblActivate1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Activating..."
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   1440
      TabIndex        =   61
      Top             =   2040
      Visible         =   0   'False
      Width           =   5175
   End
   Begin VB.Label lblReset2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "<< DISARMED >>"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   1440
      TabIndex        =   60
      Top             =   2520
      Visible         =   0   'False
      Width           =   5175
   End
   Begin VB.Label lblReset1 
      BackColor       =   &H00000000&
      Caption         =   "Disarming...Please Wait..."
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   1440
      TabIndex        =   59
      Top             =   2040
      Visible         =   0   'False
      Width           =   5175
   End
   Begin VB.Image Image30 
      Height          =   960
      Left            =   5880
      Picture         =   "frmMain.frx":0CFA
      Top             =   3960
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image Image29 
      Height          =   960
      Left            =   4920
      Picture         =   "frmMain.frx":213E
      Top             =   3960
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image Image28 
      Height          =   960
      Left            =   3960
      Picture         =   "frmMain.frx":3582
      Top             =   3960
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image Image27 
      Height          =   960
      Left            =   3000
      Picture         =   "frmMain.frx":49C6
      Top             =   3960
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image Image26 
      Height          =   960
      Left            =   2040
      Picture         =   "frmMain.frx":5E0A
      Top             =   3960
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image Image25 
      Height          =   960
      Left            =   1080
      Picture         =   "frmMain.frx":724E
      Top             =   3960
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image Image24 
      Height          =   960
      Left            =   5880
      Picture         =   "frmMain.frx":8692
      Top             =   3000
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image Image23 
      Height          =   960
      Left            =   4920
      Picture         =   "frmMain.frx":9AD6
      Top             =   3000
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image Image22 
      Height          =   960
      Left            =   3960
      Picture         =   "frmMain.frx":AF1A
      Top             =   3000
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image Image21 
      Height          =   960
      Left            =   3000
      Picture         =   "frmMain.frx":C35E
      Top             =   3000
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image Image20 
      Height          =   960
      Left            =   2040
      Picture         =   "frmMain.frx":D7A2
      Top             =   3000
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image Image19 
      Height          =   960
      Left            =   1080
      Picture         =   "frmMain.frx":EBE6
      Top             =   3000
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image Image18 
      Height          =   960
      Left            =   5880
      Picture         =   "frmMain.frx":1002A
      Top             =   2040
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image Image17 
      Height          =   960
      Left            =   4920
      Picture         =   "frmMain.frx":1146E
      Top             =   2040
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image Image16 
      Height          =   960
      Left            =   3960
      Picture         =   "frmMain.frx":128B2
      Top             =   2040
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image Image15 
      Height          =   960
      Left            =   3000
      Picture         =   "frmMain.frx":13CF6
      Top             =   2040
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image Image14 
      Height          =   960
      Left            =   2040
      Picture         =   "frmMain.frx":1513A
      Top             =   2040
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image Image13 
      Height          =   960
      Left            =   1080
      Picture         =   "frmMain.frx":1657E
      Top             =   2040
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image Image12 
      Height          =   960
      Left            =   5880
      Picture         =   "frmMain.frx":179C2
      Top             =   1080
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image Image11 
      Height          =   960
      Left            =   4920
      Picture         =   "frmMain.frx":18E06
      Top             =   1080
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image Image10 
      Height          =   960
      Left            =   3960
      Picture         =   "frmMain.frx":1A24A
      Top             =   1080
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image Image9 
      Height          =   960
      Left            =   3000
      Picture         =   "frmMain.frx":1B68E
      Top             =   1080
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image Image8 
      Height          =   960
      Left            =   2040
      Picture         =   "frmMain.frx":1CAD2
      Top             =   1080
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image Image7 
      Height          =   960
      Left            =   1080
      Picture         =   "frmMain.frx":1DF16
      Top             =   1080
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image Image6 
      Height          =   960
      Left            =   5880
      Picture         =   "frmMain.frx":1F35A
      Top             =   120
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image Image5 
      Height          =   960
      Left            =   4920
      Picture         =   "frmMain.frx":2079E
      Top             =   120
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image Image4 
      Height          =   960
      Left            =   3960
      Picture         =   "frmMain.frx":21BE2
      Top             =   120
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image Image3 
      Height          =   960
      Left            =   3000
      Picture         =   "frmMain.frx":23026
      Top             =   120
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image Image2 
      Height          =   960
      Left            =   2040
      Picture         =   "frmMain.frx":2446A
      Top             =   120
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image Image1 
      Height          =   960
      Left            =   1080
      Picture         =   "frmMain.frx":258AE
      Top             =   120
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label lblNum2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "44"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   57
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label lblNum1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "44"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   56
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label Label50 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   55
      Top             =   3840
      Width           =   495
   End
   Begin VB.Label Label49 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   54
      Top             =   3840
      Width           =   495
   End
   Begin VB.Label Label48 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   53
      Top             =   3840
      Width           =   495
   End
   Begin VB.Label Label47 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   52
      Top             =   3840
      Width           =   495
   End
   Begin VB.Label Label46 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   51
      Top             =   3840
      Width           =   495
   End
   Begin VB.Label Label45 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   50
      Top             =   3840
      Width           =   495
   End
   Begin VB.Label Label44 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   49
      Top             =   3840
      Width           =   495
   End
   Begin VB.Label Label43 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   48
      Top             =   3840
      Width           =   495
   End
   Begin VB.Label Label42 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   47
      Top             =   3840
      Width           =   495
   End
   Begin VB.Label Label41 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   46
      Top             =   3840
      Width           =   495
   End
   Begin VB.Label Label40 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   45
      Top             =   3360
      Width           =   495
   End
   Begin VB.Label Label39 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   44
      Top             =   3360
      Width           =   495
   End
   Begin VB.Label Label38 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   43
      Top             =   3360
      Width           =   495
   End
   Begin VB.Label Label37 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   42
      Top             =   3360
      Width           =   495
   End
   Begin VB.Label Label36 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   41
      Top             =   3360
      Width           =   495
   End
   Begin VB.Label Label35 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   40
      Top             =   3360
      Width           =   495
   End
   Begin VB.Label Label34 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   39
      Top             =   3360
      Width           =   495
   End
   Begin VB.Label Label33 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   38
      Top             =   3360
      Width           =   495
   End
   Begin VB.Label Label32 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   37
      Top             =   3360
      Width           =   495
   End
   Begin VB.Label Label31 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   36
      Top             =   3360
      Width           =   495
   End
   Begin VB.Label Label30 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   35
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label Label29 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   34
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   33
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label Label27 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   32
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label Label26 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   31
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label Label25 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   30
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label Label24 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   29
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   28
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   27
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   26
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   25
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   24
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   23
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   22
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   21
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   20
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   19
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   18
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   17
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   16
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   15
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   14
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   13
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   12
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   11
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   10
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   9
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   8
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   7
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   6
      Top             =   1920
      Width           =   495
   End
   Begin VB.Shape shpCountdown 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Height          =   1215
      Left            =   2400
      Top             =   1920
      Width           =   3255
   End
   Begin VB.Label lblCountdown 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   3720
      TabIndex        =   5
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label lblGetReady 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Security breach!"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   2040
      Width           =   3015
   End
   Begin VB.Label LabeliNFO 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "DETONATION IN:"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   120
      Width           =   3255
   End
   Begin VB.Shape ShapeL 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   240
      Shape           =   3  'Circle
      Top             =   840
      Width           =   615
   End
   Begin VB.Shape ShapeR 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   7200
      Shape           =   3  'Circle
      Top             =   840
      Width           =   615
   End
   Begin VB.Shape Shape4 
      BorderWidth     =   2
      Height          =   615
      Left            =   5280
      Top             =   840
      Width           =   1335
   End
   Begin VB.Shape Shape3 
      BorderWidth     =   2
      Height          =   615
      Left            =   3360
      Top             =   840
      Width           =   1335
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   2
      Height          =   615
      Left            =   1440
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label lblTIMERred 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5280
      TabIndex        =   2
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label lblTIMERyellow 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      TabIndex        =   1
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label lblTIMERgreen 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1440
      TabIndex        =   0
      Top             =   840
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   1095
      Left            =   960
      Top             =   600
      Width           =   6135
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   2535
      Left            =   960
      Top             =   1800
      Width           =   6135
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WinCount, LoseCount, SkullCount, COUNTX, NumX, Counter, Counter2, Counter3 As Integer, COUNTDOWN, TIMERred, TIMERyellow, TIMERgreen  As Integer
Private Sub Form_Load()
COUNTX = 0
SkullCount = 0
TimerWin.Enabled = False
TimerWin.Interval = 1000
WinCount = 0
lblNum1.Visible = False
lblNum2.Visible = False
Text1.Visible = False
Shape5.Visible = False
Label1.Visible = False
Label2.Visible = False
Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
Label6.Visible = False
Label7.Visible = False
Label8.Visible = False
Label9.Visible = False
Label10.Visible = False
Label11.Visible = False
Label12.Visible = False
Label13.Visible = False
Label14.Visible = False
Label15.Visible = False
Label16.Visible = False
Label17.Visible = False
Label18.Visible = False
Label19.Visible = False
Label20.Visible = False
Label21.Visible = False
Label22.Visible = False
Label23.Visible = False
Label24.Visible = False
Label25.Visible = False
Label26.Visible = False
Label27.Visible = False
Label28.Visible = False
Label29.Visible = False
Label30.Visible = False
Label31.Visible = False
Label32.Visible = False
Label33.Visible = False
Label34.Visible = False
Label35.Visible = False
Label36.Visible = False
Label37.Visible = False
Label38.Visible = False
Label39.Visible = False
Label40.Visible = False
Label41.Visible = False
Label42.Visible = False
Label43.Visible = False
Label44.Visible = False
Label45.Visible = False
Label46.Visible = False
Label47.Visible = False
Label48.Visible = False
Label49.Visible = False
Label50.Visible = False
LoseCount = 0
TimerLose.Enabled = False
TimerLose.Interval = 1000
COUNTDOWN = 3
Timer4.Interval = 1000
TIMERgreen = 100
TIMERyellow = 100
TIMERred = 100
Timer1.Interval = 250 ' Should be 250
Timer2.Interval = 250
Timer3.Interval = 250
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
TimerX.Interval = 1200
TimerSC.Enabled = False
TimerSC.Interval = 100

End Sub


Private Sub Text1_Change()
If Text1.Text = lblNum1.Caption Then COUNTX = COUNTX + 1

If COUNTX = 1 Then Label1.BackColor = &HFF00&
If COUNTX = 2 Then Label2.BackColor = &HFF00&
If COUNTX = 3 Then Label3.BackColor = &HFF00&
If COUNTX = 4 Then Label4.BackColor = &HFF00&
If COUNTX = 5 Then Label5.BackColor = &HFF00&
If COUNTX = 6 Then Label6.BackColor = &HFF00&
If COUNTX = 7 Then Label7.BackColor = &HFF00&
If COUNTX = 8 Then Label8.BackColor = &HFF00&
If COUNTX = 9 Then Label9.BackColor = &HFF00&
If COUNTX = 10 Then Label10.BackColor = &HFF00&
If COUNTX = 11 Then Label11.BackColor = &HFF00&
If COUNTX = 12 Then Label12.BackColor = &HFF00&
If COUNTX = 13 Then Label13.BackColor = &HFF00&
If COUNTX = 14 Then Label14.BackColor = &HFF00&
If COUNTX = 15 Then Label15.BackColor = &HFF00&
If COUNTX = 16 Then Label16.BackColor = &HFF00&
If COUNTX = 17 Then Label17.BackColor = &HFF00&
If COUNTX = 18 Then Label18.BackColor = &HFF00&
If COUNTX = 19 Then Label19.BackColor = &HFF00&
If COUNTX = 20 Then Label20.BackColor = &HFF00&
If COUNTX = 20 Then TimerX.Interval = 1000
If COUNTX = 21 Then Label21.BackColor = &HFF00&
If COUNTX = 22 Then Label22.BackColor = &HFF00&
If COUNTX = 23 Then Label23.BackColor = &HFF00&
If COUNTX = 24 Then Label24.BackColor = &HFF00&
If COUNTX = 25 Then Label25.BackColor = &HFF00&
If COUNTX = 26 Then Label26.BackColor = &HFF00&
If COUNTX = 27 Then Label27.BackColor = &HFF00&
If COUNTX = 28 Then Label28.BackColor = &HFF00&
If COUNTX = 29 Then Label29.BackColor = &HFF00&
If COUNTX = 30 Then Label30.BackColor = &HFF00&
If COUNTX = 31 Then Label31.BackColor = &HFF00&
If COUNTX = 32 Then Label32.BackColor = &HFF00&
If COUNTX = 33 Then Label33.BackColor = &HFF00&
If COUNTX = 34 Then Label34.BackColor = &HFF00&
If COUNTX = 35 Then Label35.BackColor = &HFF00&
If COUNTX = 36 Then Label36.BackColor = &HFF00&
If COUNTX = 37 Then Label37.BackColor = &HFF00&
If COUNTX = 38 Then Label38.BackColor = &HFF00&
If COUNTX = 39 Then Label39.BackColor = &HFF00&
If COUNTX = 40 Then Label40.BackColor = &HFF00&
If COUNTX = 40 Then TimerX.Interval = 750
If COUNTX = 41 Then Label41.BackColor = &HFF00&
If COUNTX = 42 Then Label42.BackColor = &HFF00&
If COUNTX = 43 Then Label43.BackColor = &HFF00&
If COUNTX = 44 Then Label44.BackColor = &HFF00&
If COUNTX = 45 Then Label45.BackColor = &HFF00&
If COUNTX = 45 Then TimerX.Interval = 500
If COUNTX = 46 Then Label46.BackColor = &HFF00&
If COUNTX = 47 Then Label47.BackColor = &HFF00&
If COUNTX = 48 Then Label48.BackColor = &HFF00&
If COUNTX = 49 Then Label49.BackColor = &HFF00&
If COUNTX = 50 Then Label50.BackColor = &HFF00&
If COUNTX = 50 Then
TimerWin.Enabled = True
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
TimerX.Enabled = False
lblNum1.Visible = False
lblNum2.Visible = False
Text1.Visible = False
Shape5.Visible = False
Label1.Visible = False
Label2.Visible = False
Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
Label6.Visible = False
Label7.Visible = False
Label8.Visible = False
Label9.Visible = False
Label10.Visible = False
Label11.Visible = False
Label12.Visible = False
Label13.Visible = False
Label14.Visible = False
Label15.Visible = False
Label16.Visible = False
Label17.Visible = False
Label18.Visible = False
Label19.Visible = False
Label20.Visible = False
Label21.Visible = False
Label22.Visible = False
Label23.Visible = False
Label24.Visible = False
Label25.Visible = False
Label26.Visible = False
Label27.Visible = False
Label28.Visible = False
Label29.Visible = False
Label30.Visible = False
Label31.Visible = False
Label32.Visible = False
Label33.Visible = False
Label34.Visible = False
Label35.Visible = False
Label36.Visible = False
Label37.Visible = False
Label38.Visible = False
Label39.Visible = False
Label40.Visible = False
Label41.Visible = False
Label42.Visible = False
Label43.Visible = False
Label44.Visible = False
Label45.Visible = False
Label46.Visible = False
Label47.Visible = False
Label48.Visible = False
Label49.Visible = False
Label50.Visible = False
End If

If Label1.BackColor = &HFF00& Then Label1.Caption = "X"
If Label2.BackColor = &HFF00& Then Label2.Caption = "X"
If Label3.BackColor = &HFF00& Then Label3.Caption = "X"
If Label4.BackColor = &HFF00& Then Label4.Caption = "X"
If Label5.BackColor = &HFF00& Then Label5.Caption = "X"
If Label6.BackColor = &HFF00& Then Label6.Caption = "X"
If Label7.BackColor = &HFF00& Then Label7.Caption = "X"
If Label8.BackColor = &HFF00& Then Label8.Caption = "X"
If Label9.BackColor = &HFF00& Then Label9.Caption = "X"
If Label10.BackColor = &HFF00& Then Label10.Caption = "X"
If Label11.BackColor = &HFF00& Then Label11.Caption = "X"
If Label12.BackColor = &HFF00& Then Label12.Caption = "X"
If Label13.BackColor = &HFF00& Then Label13.Caption = "X"
If Label14.BackColor = &HFF00& Then Label14.Caption = "X"
If Label15.BackColor = &HFF00& Then Label15.Caption = "X"
If Label16.BackColor = &HFF00& Then Label16.Caption = "X"
If Label17.BackColor = &HFF00& Then Label17.Caption = "X"
If Label18.BackColor = &HFF00& Then Label18.Caption = "X"
If Label19.BackColor = &HFF00& Then Label19.Caption = "X"
If Label20.BackColor = &HFF00& Then Label20.Caption = "X"
If Label21.BackColor = &HFF00& Then Label21.Caption = "X"
If Label22.BackColor = &HFF00& Then Label22.Caption = "X"
If Label23.BackColor = &HFF00& Then Label23.Caption = "X"
If Label24.BackColor = &HFF00& Then Label24.Caption = "X"
If Label25.BackColor = &HFF00& Then Label25.Caption = "X"
If Label26.BackColor = &HFF00& Then Label26.Caption = "X"
If Label27.BackColor = &HFF00& Then Label27.Caption = "X"
If Label28.BackColor = &HFF00& Then Label28.Caption = "X"
If Label29.BackColor = &HFF00& Then Label29.Caption = "X"
If Label30.BackColor = &HFF00& Then Label30.Caption = "X"
If Label31.BackColor = &HFF00& Then Label31.Caption = "X"
If Label32.BackColor = &HFF00& Then Label32.Caption = "X"
If Label33.BackColor = &HFF00& Then Label33.Caption = "X"
If Label34.BackColor = &HFF00& Then Label34.Caption = "X"
If Label35.BackColor = &HFF00& Then Label35.Caption = "X"
If Label36.BackColor = &HFF00& Then Label36.Caption = "X"
If Label37.BackColor = &HFF00& Then Label37.Caption = "X"
If Label38.BackColor = &HFF00& Then Label38.Caption = "X"
If Label39.BackColor = &HFF00& Then Label39.Caption = "X"
If Label40.BackColor = &HFF00& Then Label40.Caption = "X"
If Label41.BackColor = &HFF00& Then Label41.Caption = "X"
If Label42.BackColor = &HFF00& Then Label42.Caption = "X"
If Label43.BackColor = &HFF00& Then Label43.Caption = "X"
If Label44.BackColor = &HFF00& Then Label44.Caption = "X"
If Label45.BackColor = &HFF00& Then Label45.Caption = "X"
If Label46.BackColor = &HFF00& Then Label46.Caption = "X"
If Label47.BackColor = &HFF00& Then Label47.Caption = "X"
If Label48.BackColor = &HFF00& Then Label48.Caption = "X"
If Label49.BackColor = &HFF00& Then Label49.Caption = "X"
If Label50.BackColor = &HFF00& Then Label50.Caption = "X"
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyReturn: Text1.Text = ""
End Select
End Sub

Private Sub Timer1_Timer()
TIMERgreen = TIMERgreen - 1
lblTIMERgreen = TIMERgreen
If TIMERgreen = 0 Then
ShapeL.FillColor = &HFFFF&
ShapeR.FillColor = &HFFFF&
Timer1.Enabled = False
Timer2.Enabled = True
End If
End Sub

Private Sub Timer2_Timer()
TIMERyellow = TIMERyellow - 1
lblTIMERyellow = TIMERyellow
If TIMERyellow = 0 Then
ShapeL.FillColor = &HFF&
ShapeR.FillColor = &HFF&
Timer2.Enabled = False
Timer3.Enabled = True
End If
End Sub

Private Sub Timer3_Timer()

TIMERred = TIMERred - 1
lblTIMERred = TIMERred
If TIMERred = 0 Then
TimerLose.Enabled = True
frmMain.BackColor = &H0&
Shape1.Visible = False
lblTIMERred.Visible = False
lblTIMERyellow.Visible = False
lblTIMERgreen.Visible = False
LabeliNFO.Visible = False
Timer1.Enabled = False
ShapeL.Visible = False
ShapeR.Visible = False
Timer2.Enabled = False
Timer3.Enabled = False
TimerX.Enabled = False
lblNum1.Visible = False
lblNum2.Visible = False
Text1.Visible = False
Shape5.Visible = False
Label1.Visible = False
Label2.Visible = False
Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
Label6.Visible = False
Label7.Visible = False
Label8.Visible = False
Label9.Visible = False
Label10.Visible = False
Label11.Visible = False
Label12.Visible = False
Label13.Visible = False
Label14.Visible = False
Label15.Visible = False
Label16.Visible = False
Label17.Visible = False
Label18.Visible = False
Label19.Visible = False
Label20.Visible = False
Label21.Visible = False
Label22.Visible = False
Label23.Visible = False
Label24.Visible = False
Label25.Visible = False
Label26.Visible = False
Label27.Visible = False
Label28.Visible = False
Label29.Visible = False
Label30.Visible = False
Label31.Visible = False
Label32.Visible = False
Label33.Visible = False
Label34.Visible = False
Label35.Visible = False
Label36.Visible = False
Label37.Visible = False
Label38.Visible = False
Label39.Visible = False
Label40.Visible = False
Label41.Visible = False
Label42.Visible = False
Label43.Visible = False
Label44.Visible = False
Label45.Visible = False
Label46.Visible = False
Label47.Visible = False
Label48.Visible = False
Label49.Visible = False
Label50.Visible = False
Timer3.Enabled = False
TimerSC.Enabled = True
End If
End Sub

Private Sub Timer4_Timer()
COUNTDOWN = COUNTDOWN - 1
lblCountdown = COUNTDOWN
If COUNTDOWN = 0 Then
lblNum1.Visible = True
lblNum2.Visible = True
Text1.Visible = True
Shape5.Visible = True
Label1.Visible = True
Label2.Visible = True
Label3.Visible = True
Label4.Visible = True
Label5.Visible = True
Label6.Visible = True
Label7.Visible = True
Label8.Visible = True
Label9.Visible = True
Label10.Visible = True
Label11.Visible = True
Label12.Visible = True
Label13.Visible = True
Label14.Visible = True
Label15.Visible = True
Label16.Visible = True
Label17.Visible = True
Label18.Visible = True
Label19.Visible = True
Label20.Visible = True
Label21.Visible = True
Label22.Visible = True
Label23.Visible = True
Label24.Visible = True
Label25.Visible = True
Label26.Visible = True
Label27.Visible = True
Label28.Visible = True
Label29.Visible = True
Label30.Visible = True
Label31.Visible = True
Label32.Visible = True
Label33.Visible = True
Label34.Visible = True
Label35.Visible = True
Label36.Visible = True
Label37.Visible = True
Label38.Visible = True
Label39.Visible = True
Label40.Visible = True
Label41.Visible = True
Label42.Visible = True
Label43.Visible = True
Label44.Visible = True
Label45.Visible = True
Label46.Visible = True
Label47.Visible = True
Label48.Visible = True
Label49.Visible = True
Label50.Visible = True
Timer4.Enabled = False
shpCountdown.Visible = False
lblCountdown.Visible = False
lblGetReady.Visible = False
Timer1.Enabled = True
End If
End Sub

Private Sub TimerLose_Timer()
LoseCount = LoseCount + 1
If LoseCount = 1 Then lblActivate1.Visible = True
If LoseCount = 3 Then lblActivate2.Visible = True
If LoseCount = 4 Then lblActivate2.Visible = False
If LoseCount = 5 Then lblActivate2.Visible = True
If LoseCount = 6 Then lblActivate2.Visible = False
If LoseCount = 7 Then lblActivate2.Visible = True
If LoseCount = 9 Then
Unload Me
frmLose.Show
End If
End Sub

Private Sub TimerWin_Timer()
WinCount = WinCount + 1
If WinCount = 1 Then lblReset1.Visible = True
If WinCount = 3 Then lblReset2.Visible = True
If WinCount = 5 Then
Unload Me
frmWin.Show
End If
End Sub

Private Sub TimerX_Timer()
NumX = Int(50 * Rnd + 1)
lblNum1 = NumX
lblNum2 = NumX
End Sub



Private Sub TimerSC_Timer()
SkullCount = SkullCount + 1

If SkullCount = 1 Then Image1.Visible = True
If SkullCount = 2 Then Image2.Visible = True
If SkullCount = 3 Then Image3.Visible = True
If SkullCount = 4 Then Image4.Visible = True
If SkullCount = 5 Then Image5.Visible = True
If SkullCount = 6 Then Image6.Visible = True
If SkullCount = 7 Then Image12.Visible = True
If SkullCount = 8 Then Image24.Visible = True
If SkullCount = 9 Then Image30.Visible = True
If SkullCount = 10 Then Image29.Visible = True
If SkullCount = 11 Then Image28.Visible = True
If SkullCount = 12 Then Image27.Visible = True
If SkullCount = 13 Then Image26.Visible = True
If SkullCount = 14 Then Image25.Visible = True
If SkullCount = 15 Then Image19.Visible = True
If SkullCount = 16 Then Image7.Visible = True
If SkullCount = 17 Then Image8.Visible = True
If SkullCount = 18 Then Image9.Visible = True
If SkullCount = 19 Then Image10.Visible = True
If SkullCount = 20 Then Image11.Visible = True
If SkullCount = 21 Then Image23.Visible = True
If SkullCount = 22 Then Image22.Visible = True
If SkullCount = 23 Then Image21.Visible = True
If SkullCount = 24 Then Image20.Visible = True
End Sub

