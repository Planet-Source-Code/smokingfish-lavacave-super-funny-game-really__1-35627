VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Lava Cave"
   ClientHeight    =   2685
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3420
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000E&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   3420
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      Height          =   495
      Left            =   480
      ScaleHeight     =   435
      ScaleWidth      =   315
      TabIndex        =   28
      Top             =   1440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   615
      Left            =   0
      Picture         =   "frmMain.frx":0E42
      ScaleHeight     =   555
      ScaleWidth      =   435
      TabIndex        =   27
      Top             =   1440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox picPlayermask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   2160
      Picture         =   "frmMain.frx":3462
      ScaleHeight     =   180
      ScaleWidth      =   1080
      TabIndex        =   25
      Top             =   1320
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.ListBox lstY 
      Height          =   300
      Left            =   720
      TabIndex        =   24
      Top             =   2160
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ListBox lstX 
      Height          =   300
      Left            =   0
      TabIndex        =   23
      Top             =   2160
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Timer Timer3 
      Interval        =   150
      Left            =   2640
      Top             =   360
   End
   Begin VB.PictureBox pic200mask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   0
      Picture         =   "frmMain.frx":3EC4
      ScaleHeight     =   240
      ScaleWidth      =   300
      TabIndex        =   22
      Top             =   2520
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox pic200pmask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   0
      Picture         =   "frmMain.frx":3F1E
      ScaleHeight     =   240
      ScaleWidth      =   300
      TabIndex        =   21
      Top             =   2880
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox pic500pmask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   0
      Picture         =   "frmMain.frx":3F78
      ScaleHeight     =   240
      ScaleWidth      =   300
      TabIndex        =   20
      Top             =   3240
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picspeedmask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   720
      Picture         =   "frmMain.frx":3FD6
      ScaleHeight     =   180
      ScaleWidth      =   300
      TabIndex        =   19
      Top             =   3600
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picwallsmask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   720
      Picture         =   "frmMain.frx":401B
      ScaleHeight     =   240
      ScaleWidth      =   300
      TabIndex        =   18
      Top             =   3240
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picenemyslowmask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   720
      Picture         =   "frmMain.frx":4361
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   17
      Top             =   2520
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picspeedslowmask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   720
      Picture         =   "frmMain.frx":43B3
      ScaleHeight     =   240
      ScaleWidth      =   300
      TabIndex        =   16
      Top             =   2880
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picwallbackmask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   0
      Picture         =   "frmMain.frx":4422
      ScaleHeight     =   240
      ScaleWidth      =   300
      TabIndex        =   15
      Top             =   3600
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox pickillmask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   0
      Picture         =   "frmMain.frx":4768
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   14
      Top             =   3960
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox pickill 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   360
      Picture         =   "frmMain.frx":47D7
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   13
      Top             =   3960
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picenemy 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   2160
      Picture         =   "frmMain.frx":4876
      ScaleHeight     =   180
      ScaleWidth      =   720
      TabIndex        =   12
      Top             =   1080
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.PictureBox picwallback 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   360
      Picture         =   "frmMain.frx":4B2A
      ScaleHeight     =   240
      ScaleWidth      =   300
      TabIndex        =   11
      Top             =   3600
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picspeedslow 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1080
      Picture         =   "frmMain.frx":4BF6
      ScaleHeight     =   240
      ScaleWidth      =   300
      TabIndex        =   10
      Top             =   2880
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picenemyslow 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1080
      Picture         =   "frmMain.frx":4CFD
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   9
      Top             =   2520
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picWalls 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1080
      Picture         =   "frmMain.frx":4DB3
      ScaleHeight     =   240
      ScaleWidth      =   300
      TabIndex        =   8
      Top             =   3240
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picSpeed 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   1080
      Picture         =   "frmMain.frx":4E8E
      ScaleHeight     =   180
      ScaleWidth      =   300
      TabIndex        =   7
      Top             =   3600
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox pic500p 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   360
      Picture         =   "frmMain.frx":4FA3
      ScaleHeight     =   240
      ScaleWidth      =   300
      TabIndex        =   6
      Top             =   3240
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox pic200p 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   360
      Picture         =   "frmMain.frx":50ED
      ScaleHeight     =   240
      ScaleWidth      =   300
      TabIndex        =   5
      Top             =   2880
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox pic200 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   360
      Picture         =   "frmMain.frx":5231
      ScaleHeight     =   240
      ScaleWidth      =   300
      TabIndex        =   4
      Top             =   2520
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picPlayer 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   2160
      Picture         =   "frmMain.frx":5324
      ScaleHeight     =   180
      ScaleWidth      =   1080
      TabIndex        =   3
      Top             =   840
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   2640
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   400
      Left            =   3000
      Top             =   0
   End
   Begin VB.Label Label4 
      BackColor       =   &H00404040&
      Caption         =   "-?-"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   120
      Width           =   280
   End
   Begin VB.Label Label3 
      BackColor       =   &H00404040&
      Caption         =   "End Game"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2040
      TabIndex        =   2
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00404040&
      Caption         =   "About"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2040
      TabIndex        =   1
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404040&
      Caption         =   "Start Game"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2040
      TabIndex        =   0
      Top             =   1560
      Width           =   1215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim UpDown As Boolean
Dim EnemyTrue1 As Boolean
Dim EnemyTrue2 As Boolean
Dim EnemyTrue3 As Boolean
Dim Score1 As Integer
Dim Score2 As Integer

Private Sub Form_Load()
Me.Picture = Picture1.Picture
Label1.Visible = True
Label2.Visible = True
Label3.Visible = True
'Label4.Visible = True
Game.EnemyFrame = 0
Game.SideMove = 0
Game.SnakeAcceleration = 0
Game.SnakeOver = False
Game.SnakePositionX = 0
Game.SnakePositionY = 0
Game.SnakeSpeed = 0
Game.WorldGravity = 0
Game.WorldWidth = 0
EnemyTrue1 = False
EnemyTrue2 = False
EnemyTrue3 = False
UpDown = False
Score1 = 0
Score2 = 0
Me.Show
End Sub

Public Sub MainLoop()
Dim TEMPa As Integer
Dim TEMPb As Integer
Do
DoEvents
Me.Cls
Sleep 5
Game.SnakeSpeed = Game.SnakeSpeed + Game.WorldGravity

If Game.SnakePositionY + Game.SnakeSpeed + 180 > 2500 Then
    Game.SnakeSpeed = -Int(Game.SnakeSpeed * 0.5)
    Game.SnakePositionY = 2500 - 180
Else
    Game.SnakePositionY = Game.SnakePositionY + Game.SnakeSpeed
End If

If Game.SnakeDirection = True Then
    If Game.SnakePositionX + 230 + Game.SnakeAcceleration > 3200 Then
        Game.SnakeDirection = False
        Game.SnakePositionX = 3200 - 230
        Game.SnakeAcceleration = Game.SnakeAcceleration + 4
    Else
    Game.SnakePositionX = Game.SnakePositionX + Game.SnakeAcceleration
    End If
End If

If Game.SnakeDirection = False Then
    If Game.SnakePositionX - Game.SnakeAcceleration < 254 Then
        Game.SnakeDirection = True
        Game.SnakePositionX = 254
        Game.SnakeAcceleration = Game.SnakeAcceleration + 4
    Else
        Game.SnakePositionX = Game.SnakePositionX - Game.SnakeAcceleration
    End If
End If

If Game.SnakePositionY <= 254 Then
Game.SnakePositionY = 254
Game.SnakeSpeed = -Int(Game.SnakeSpeed * 0.5)
End If

If GetAsyncKeyState(vbKeySpace) Then
Sleep 5
Game.SnakeSpeed = Game.SnakeSpeed - 4
End If

Game.DrawEnemys
Game.DrawWorld
Game.CheckEnemys
Game.DrawSnake

If Not frmMain.Point(Game.SnakePositionX, Game.SnakePositionY) = vbBlack Then
Game.SnakeOver = True
End If
If Not frmMain.Point(Game.SnakePositionX + 170, Game.SnakePositionY) = vbBlack Then
Game.SnakeOver = True
End If
If Not frmMain.Point(Game.SnakePositionX, Game.SnakePositionY + 170) = vbBlack Then
Game.SnakeOver = True
End If
If Not frmMain.Point(Game.SnakePositionX + 170, Game.SnakePositionY + 170) = vbBlack Then
Game.SnakeOver = True
End If
 
'frmMain.PSet (Game.SnakePositionX, Game.SnakePositionY)
'frmMain.PSet (Game.SnakePositionX + 170, Game.SnakePositionY)
'frmMain.PSet (Game.SnakePositionX, Game.SnakePositionY + 170)
'frmMain.PSet (Game.SnakePositionX + 170, Game.SnakePositionY + 170)

Game.DrawSnakeShadow

Score1 = Score1 + 1
frmMain.CurrentX = 1100
frmMain.CurrentY = 2200
frmMain.Print "Score: " & Score1
frmMain.CurrentX = 900
frmMain.CurrentY = 0
frmMain.Print "Highscore: " & Score2


Loop Until Game.SnakeOver = True
frmMain.Cls
Game.DrawWorld
For i = 1 To 100
Sleep 5
Game.GradientCircle frmMain, Game.SnakePositionX, Game.SnakePositionY, i * 5, 200, 50, 50, 5, False, False
frmMain.Refresh
Next i
Sleep 1000
frmMain.Cls
frmMain.CurrentX = 1100
frmMain.CurrentY = 500
frmMain.Print "Game Over!!"
frmMain.CurrentX = 1100
frmMain.CurrentY = 2200
frmMain.Print "Score: " & Score1
If Score2 < Score1 Then
Open App.Path & "\Score.Dat" For Binary As #1
Put #1, , Score1
Close #1
MsgBox "NEW HIGHSCORE!!!"
End If
Me.Refresh
Sleep 3000
'Shell App.Path & "\" & App.EXEName & ".exe", vbNormalFocus
Me.Cls
Form_Load
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Label1_Click()
Me.Picture = Picture2.Picture
Label1.Visible = False
Label2.Visible = False
Label3.Visible = False
Label4.Visible = False
Score1 = 0
Score2 = 0
If FileExists(App.Path & "\Score.Dat") = True Then
Open App.Path & "\Score.Dat" For Binary As #1
Get #1, , Score2
Close #1
End If
Game.SnakePositionX = 400
Game.SnakePositionY = 1250
Game.SnakeSpeed = 0
Game.SnakeDirection = True
Game.WorldGravity = 1.6
Game.DummiWorld
Game.SnakeAcceleration = 10
Game.SideMove = 25
Game.WorldWidth = 150
UpDown = True
Call MainLoop
End Sub

Private Sub Label2_Click()
MsgBox "2002 by SmokingFish!" & vbCrLf & "mail@SmokingFish.de", vbInformation
End Sub

Private Sub Label3_Click()
End
End Sub

Private Sub Label4_Click()
MsgBox "Use the Space Key to accelerate! Dont touch the Red and the Blue Wall or the Enemys ! And remember ! This Game is Truecolor only!", vbInformation
End Sub

Private Sub Timer1_Timer()
If UpDown = True Then
Game.WorldWidth = Game.WorldWidth - 1
Else
Game.WorldWidth = Game.WorldWidth + 1
End If
If Game.WorldWidth = 80 Then UpDown = False
If Game.WorldWidth = 110 Then UpDown = True
End Sub

Private Sub Timer3_Timer()
lstX.AddItem Game.SnakePositionX
lstY.AddItem Game.SnakePositionY
If lstX.ListCount = 6 Then
lstX.RemoveItem 0
lstY.RemoveItem 0
End If
If Game.EnemyFrame = 0 Then
Game.EnemyFrame = 1
Exit Sub
End If
If Game.EnemyFrame = 1 Then
Game.EnemyFrame = 2
Exit Sub
End If
If Game.EnemyFrame = 2 Then
Game.EnemyFrame = 3
Exit Sub
End If
If Game.EnemyFrame = 3 Then
Game.EnemyFrame = 0
Exit Sub
End If
End Sub
