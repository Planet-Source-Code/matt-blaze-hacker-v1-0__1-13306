VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3720
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7185
   LinkTopic       =   "Form1"
   ScaleHeight     =   3720
   ScaleWidth      =   7185
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer3 
      Left            =   5400
      Top             =   1080
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4800
      TabIndex        =   6
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Timer Timer2 
      Left            =   3240
      Top             =   1080
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2640
      TabIndex        =   4
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Left            =   1080
      Top             =   1080
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   480
      TabIndex        =   0
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label Label6 
      Height          =   495
      Left            =   4800
      TabIndex        =   8
      Top             =   360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label5 
      Height          =   495
      Left            =   2640
      TabIndex        =   7
      Top             =   360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   5
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   1680
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Counter, Counter2, Counter3 As Integer
Private Sub Form_Load()
Counter = 0
Timer1.Interval = 200
Counter2 = 0
Timer2.Interval = 200
Counter3 = 0
Timer3.Interval = 200
End Sub

Private Sub Text1_Change()
If Text1.Text = "THIS IS A TEST TO SEE HOW QUICK YOUR REFLEXES ARE" Then
Text1.BackColor = &HFF00&
Label1.Caption = "PASSED"
Timer1.Enabled = False
End If
End Sub

Private Sub Text2_Change()
If Text2.Text = "PETER PIPER PICKED A PECK OF PICKLED PEPPER WHERES THE PECK OF PICKLED PEPPER PETER PIPER PICKED?" Then
Text2.BackColor = &HFF00&
Label3.Caption = "PASSED"
Timer2.Enabled = False
End If
End Sub

Private Sub Text3_Change()
If Text3.Text = "30" Then
Text3.BackColor = &HFF00&
Label4.Caption = "PASSED"
Timer3.Enabled = False
End If
End Sub

Private Sub Timer1_Timer()
Counter = Counter + 1
Label2 = Counter
If Counter = 1 Then Label1.Caption = "THIS"
If Counter = 2 Then Label1.Caption = "IS"
If Counter = 3 Then Label1.Caption = "A"
If Counter = 4 Then Label1.Caption = "TEST"
If Counter = 5 Then Label1.Caption = "TO"
If Counter = 6 Then Label1.Caption = "SEE"
If Counter = 7 Then Label1.Caption = "HOW"
If Counter = 8 Then Label1.Caption = "QUICK"
If Counter = 9 Then Label1.Caption = "YOUR"
If Counter = 10 Then Label1.Caption = "REFLEXES"
If Counter = 11 Then Label1.Caption = "ARE"
If Counter = 13 Then Counter = 0

End Sub

Private Sub Timer2_Timer()
Counter2 = Counter2 + 1
Label5 = Counter2
If Counter2 = 1 Then Label3.Caption = "PETER"
If Counter2 = 2 Then Label3.Caption = "PIPER"
If Counter2 = 3 Then Label3.Caption = "PICKED"
If Counter2 = 4 Then Label3.Caption = "A"
If Counter2 = 5 Then Label3.Caption = "PECK"
If Counter2 = 6 Then Label3.Caption = "OF"
If Counter2 = 7 Then Label3.Caption = "PICKLED"
If Counter2 = 8 Then Label3.Caption = "PEPPER"
If Counter2 = 9 Then Label3.Caption = "WHERES"
If Counter2 = 10 Then Label3.Caption = "THE"
If Counter2 = 11 Then Label3.Caption = "PECK"
If Counter2 = 12 Then Label3.Caption = "OF"
If Counter2 = 13 Then Label3.Caption = "PICKLED"
If Counter2 = 14 Then Label3.Caption = "PEPPER"
If Counter2 = 15 Then Label3.Caption = "PETER"
If Counter2 = 16 Then Label3.Caption = "PIPER"
If Counter2 = 17 Then Label3.Caption = "PICKED"
If Counter2 = 18 Then Label3.Caption = "?"
If Counter2 = 20 Then Counter2 = 0
End Sub

Private Sub Timer3_Timer()
Counter3 = Counter3 + 1
Label6 = Counter3
If Counter3 = 1 Then Label4.Caption = "3"
If Counter3 = 2 Then Label4.Caption = "x"
If Counter3 = 3 Then Label4.Caption = "7"
If Counter3 = 4 Then Label4.Caption = "+"
If Counter3 = 5 Then Label4.Caption = "10"
If Counter3 = 6 Then Label4.Caption = "-"
If Counter3 = 7 Then Label4.Caption = "1"
If Counter3 = 8 Then Label4.Caption = "="
If Counter3 = 9 Then Label4.Caption = "?"
If Counter3 = 11 Then Counter3 = 0
End Sub
