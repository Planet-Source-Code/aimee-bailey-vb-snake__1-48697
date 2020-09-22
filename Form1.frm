VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00AD844A&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "VB-Snake (By Steve B)"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5910
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   5910
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C6B6AA&
      Caption         =   "&Minimize"
      Height          =   375
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C6B6AA&
      Caption         =   "&Stop"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5280
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C6B6AA&
      Caption         =   "&About"
      Height          =   375
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C6B6AA&
      Caption         =   "&Exit"
      Height          =   375
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C6B6AA&
      Caption         =   "&New Game"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5280
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   4395
      Left            =   95
      MousePointer    =   2  'Cross
      Picture         =   "Form1.frx":0442
      ScaleHeight     =   4335
      ScaleWidth      =   5655
      TabIndex        =   0
      Top             =   720
      Width           =   5720
      Begin VB.PictureBox wall 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   135
         Index           =   0
         Left            =   1560
         Picture         =   "Form1.frx":59EF8
         ScaleHeight     =   135
         ScaleWidth      =   135
         TabIndex        =   14
         Top             =   3480
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.PictureBox Shape3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   135
         Left            =   1800
         Picture         =   "Form1.frx":5A036
         ScaleHeight     =   135
         ScaleWidth      =   135
         TabIndex        =   12
         Top             =   3480
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Timer Timer4 
         Interval        =   10
         Left            =   1560
         Top             =   3840
      End
      Begin VB.Timer Timer3 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   1080
         Top             =   3840
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   600
         Left            =   600
         Top             =   3840
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   300
         Left            =   120
         Top             =   3840
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Paused."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   11
         Top             =   120
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   72
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   1695
         Left            =   0
         TabIndex        =   3
         Top             =   1200
         Width           =   5535
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         Height          =   135
         Left            =   2040
         Shape           =   3  'Circle
         Top             =   3480
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H0000C000&
         BackStyle       =   1  'Opaque
         Height          =   135
         Index           =   0
         Left            =   480
         Shape           =   3  'Circle
         Top             =   120
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H0000C000&
         BackStyle       =   1  'Opaque
         Height          =   135
         Index           =   1
         Left            =   600
         Shape           =   3  'Circle
         Top             =   120
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H0000C000&
         BackStyle       =   1  'Opaque
         Height          =   135
         Index           =   2
         Left            =   720
         Shape           =   3  'Circle
         Top             =   120
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H0000C000&
         BackStyle       =   1  'Opaque
         Height          =   135
         Index           =   3
         Left            =   840
         Shape           =   3  'Circle
         Top             =   120
         Visible         =   0   'False
         Width           =   135
      End
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00AD844A&
      Caption         =   "Crazy Colorz"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   5280
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C6B6AA&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   13
      Top             =   360
      Width           =   3615
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C6B6AA&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   9
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Score"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public xMove As Long
Public yMove As Long
Public Movement As Integer
Public snakeLength As Integer
Public LastIndex As Integer
Public SmallScore As Integer
Public xtraScore As Integer
Public ShownEnd As Boolean
Public Player As String

Public colSnake As Long
Public colBack As Long
Public colGoal As Long
Public colXGoal As Long
Public HighestScore As Integer
Public HighestPlayer As String
Public xset As Integer
Public xt As Integer

Public Function LoadSettings()
On Error GoTo err
Open App.Path & "\users.dat" For Input As #1
Input #1, HighestScore
Input #1, HighestPlayer
err:
Close #1
End Function

Public Function SaveSettings()
On Error Resume Next
Open App.Path & "\users.dat" For Output As #2
Print #2, HighestScore
Print #2, HighestPlayer
Close #2
End Function

Public Function RandomColor() As Long
Dim r, g, b As Integer
Randomize
r = Int(Rnd * 255)
Randomize
g = Int(Rnd * 255)
Randomize
b = Int(Rnd * 255)
If r < 100 Then r = 100
If g < 100 Then g = 100
If b < 100 Then b = 100
RandomColor = RGB(r, g, b)
End Function


Public Function CaughtGoal()
Dim xx As Long
Label3.Caption = Int(Label3.Caption) + Int(SmallScore)
NewGoal: NewGoal
Gulp
'NEW
If Check1.Value = 1 Then
    xx = RandomColor
    For a = 0 To Shape1.Count - 1
        Shape1(a).BackColor = xx
    Next a
    colBack = RandomColor
    Shape2.BackColor = RandomColor
    Picture1.BackColor = RandomColor
End If

xx = Shape1.Count
Load Shape1(xx)
Shape1(xx).Visible = True
End Function
Public Function Gulp()
beep
End Function
Public Function Tada()
beep: beep
End Function
Public Function CaughtxGoal()
Dim xx As Long
Tada
Label3.Caption = Int(Label3.Caption) + Int(xtraScore)
xtraScore = xtraScore + SmallScore
Newxgoal: Newxgoal

'NEW
If Check1.Value = 1 Then
    xx = RandomColor
    For a = 0 To Shape1.Count - 1
        Shape1(a).BackColor = xx
    Next a
    colBack = RandomColor
    Shape2.BackColor = RandomColor
    Picture1.BackColor = RandomColor
End If

xx = Shape1.Count
Load Shape1(xx)
Shape1(xx).Visible = True
End Function
Public Function ResetGame()
On Error Resume Next
Timer1.Enabled = False

If Shape1.Count > 3 Then
    For j = 4 To Shape1.Count
        Unload Shape1(j)
    Next j
End If

Shape1(0).Top = 120
For i = 0 To 3
 Shape1(i).Left = Int(i * 120) + 240
 Shape1(i).Top = 120
 Shape1(i).Visible = False
Next i

Shape2.Visible = False

Timer4.Enabled = True
End Function
Public Function NewGoal()
Dim x, y As Integer
Dim maxx, maxy As Integer

Shape2.Visible = False
redo:
      DoEvents
        maxx = Int(GetMaxHeight / 120)
        maxy = Int(GetMaxHeight / 120)
        Randomize
        x = Int(Rnd * maxx) * 120
        Randomize
        y = Int(Rnd * maxy) * 120
    
If x < 0 Then GoTo redo
If y < 0 Then GoTo redo
If y > GetMaxHeight Then GoTo redo
If x > GetMaxWidth Then GoTo redo
If x < 120 Then x = 120
If y < 120 Then y = 120
If x > (GetMaxWidth - 120) Then x = (GetMaxWidth - 120)
If y > (GetMaxHeight - 120) Then y = (GetMaxHeight - 120)


Shape2.Left = x
Shape2.Top = y
Shape2.Visible = True

End Function
Public Function Newxgoal()
Dim x, y As Integer
Dim maxx, maxy As Integer
Shape3.Visible = False
redo:
      DoEvents
        maxx = Int(GetMaxHeight / 120)
        maxy = Int(GetMaxHeight / 120)
        Randomize
        x = Int(Rnd * maxx) * 120
        Randomize
        y = Int(Rnd * maxy) * 120
    
If x < 0 Then GoTo redo
If y < 0 Then GoTo redo
If y > GetMaxHeight Then GoTo redo
If x > GetMaxWidth Then GoTo redo

Shape3.Left = x
Shape3.Top = y
Shape3.Visible = True
End Function

Private Sub Check1_KeyPress(KeyAscii As Integer)
Picture1_KeyPress KeyAscii
End Sub

Private Sub Command1_Click()
ResetGame
snakeLength = 3
Movement = 3
xMove = 120
ShownEnd = False
Form2.Show vbModal, Me
End Sub

Private Sub Command1_KeyPress(KeyAscii As Integer)
Picture1_KeyPress KeyAscii
End Sub

Private Sub Command2_Click()
x = MsgBox("Are you sure?", vbYesNo, "VB Snake")
If x = vbYes Then End
End Sub

Private Sub Command2_KeyPress(KeyAscii As Integer)
Picture1_KeyPress KeyAscii
End Sub

Private Sub Command3_Click()
frmAbout.Show vbModal, Me
End Sub

Private Sub Command3_KeyPress(KeyAscii As Integer)
Picture1_KeyPress KeyAscii
End Sub

Private Sub Command4_Click()
Timer1.Enabled = False
Timer4.Enabled = False
Command4.Visible = False
End Sub

Private Sub Command5_Click()
Me.WindowState = 1
End Sub

Private Sub Command5_KeyPress(KeyAscii As Integer)
Picture1_KeyPress KeyAscii
End Sub

Private Sub form_KeyPress(KeyAscii As Integer)
Picture1_KeyPress KeyAscii
End Sub
Private Sub Form_Load()
On Error Resume Next
Dim xx As String
If Right(App.Path, 1) <> "\" Then xx = "\"
Picture1.Picture = LoadPicture(App.Path & xx & "grid.bmp")
Form1.Command4.Visible = False
snakeLength = 3
Movement = 3
xMove = 120
xtraScore = 10

If Check1.Value = 1 Then
    xx = RandomColor
    For a = 0 To Shape1.Count - 1
        Shape1(a).BackColor = xx
    Next a
    colBack = RandomColor
    Shape2.BackColor = RandomColor
    Picture1.BackColor = RandomColor
End If

LoadSettings
Label4.Caption = "Current Highest: " & HighestPlayer & " with " & HighestScore

Form2.Show vbModal, Me
End Sub
Public Function ClearScr()
For i = 0 To Shape1.Count - 1
Shape1(i).Visible = False
Next i
Shape2.Visible = False
Shape3.Visible = False
End Function
Private Sub Picture1_Click()
NewGoal
End Sub

Private Sub Picture1_KeyPress(KeyAscii As Integer)
oldXmove = xMove
oldYmove = yMove

If Label5.Visible Then
    If KeyAscii <> 53 Then Exit Sub
End If

Select Case KeyAscii
    Case 53
        If Label5.Visible = True Then
            Timer1.Enabled = True
            Label5.Visible = False
        Else
            Timer1.Enabled = False
            Label5.Visible = True
        End If
    Case 56
        If yMove = 120 Then GoTo err
        xMove = 0: yMove = -120 'up
    Case 50
        If yMove = -120 Then GoTo err
        xMove = 0: yMove = 120 'down
    Case 54
        If xMove = -120 Then GoTo err
        xMove = 120: yMove = 0 'left
    Case 52
        If xMove = 120 Then GoTo err
        xMove = -120: yMove = 0 'right
End Select
err:
End Sub
Private Sub Timer1_Timer()
On Error Resume Next
Timer2.Enabled = False
snakeLength = Shape1.Count - 1
Movement = Movement + 1
If Movement = snakeLength + 1 Then Movement = 0

If Movement = 0 Then
LastIndex = snakeLength
Else
LastIndex = Int(Movement - 1)
End If

Shape1(Movement).Shape = 3
Shape1(Movement).Left = Shape1(LastIndex).Left + xMove
Shape1(Movement).Top = Shape1(LastIndex).Top + yMove
Shape1(Movement).Visible = True

'#### Never Ending Walls!
If Shape1(Movement).Top < 0 Then Shape1(Movement).Top = GetMaxHeight
If Shape1(Movement).Left < 0 Then Shape1(Movement).Left = GetMaxWidth
If Shape1(Movement).Top > GetMaxHeight Then Shape1(Movement).Top = 0
If Shape1(Movement).Left > GetMaxWidth Then Shape1(Movement).Left = 0
'#### Never Ending Walls!
Collide
If Shape1(Movement).Left = Shape2.Left Then
    If Shape1(Movement).Top = Shape2.Top Then
        CaughtGoal
        Shape1(Movement).Shape = 1
        xt = xt + 1
        If xset = 0 Then Label6.Caption = "Oh Well!"
        If xt = 5 Then
            xt = 0
            xset = 80
            Newxgoal
        End If
    End If
End If

If Shape1(Movement).Left = Shape3.Left Then
    If Shape1(Movement).Top = Shape3.Top Then
        CaughtxGoal
        Shape1(Movement).Shape = 1
        Shape3.Visible = False
        xset = 0
    End If
End If

If xset > 0 Then
    Label6.Caption = "Hurry And Get The Bonus!! " & xset
    xset = xset - 1
ElseIf xset = 0 Then
    Shape3.Visible = False
    Shape3.Left = Me.Width
    Label6.Caption = (HighestScore - Label3.Caption + 1) & " Left to beat the leader!"
End If

End Sub
Public Function wallColide()
For i = 1 To wall.Count - 1
If Shape1(Movement).Left = wall(i).Left Then
    If Shape1(Movement).Top = wall(i).Top Then
        DoColide
    End If
End If
Next i
End Function
Public Function DoColide()
 Timer1.Enabled = False
          Timer4.Enabled = False
          ShownEnd = True
          MsgBox "Game Over!" & vbCrLf & "---------------" & vbCrLf & "You Scored " & Label3.Caption & "!!!", vbOKOnly, "VB Snake"
          
          If HighestScore < Label3.Caption Then
            MsgBox "Guess What! You beat " & HighestPlayer & " by " & Label3.Caption & "/" & HighestScore
            Me.HighestScore = Label3.Caption
            Me.HighestPlayer = Me.Player
            SaveSettings
            Label4.Caption = "Current Highest: " & HighestPlayer & " with " & HighestScore
          Else
            If Int(Label3.Caption) > (HighestScore - 10) Then
                MsgBox "You were so close to beating him/her!, please try again!", vbOKOnly, "VB-Snake"
            Else
                MsgBox "Are you going to let " & HighestPlayer & " get away with winning all the time?" & vbCrLf & "if not, try to beat his/her score of " & HighestScore & "!!!", vbOKOnly, "VB-Snake"
            End If
            Command4.Visible = False
          End If
aa:
          ResetGame
          xxx = MsgBox("Play again?", vbYesNo, "VB-Snake")
          If xxx = vbYes Then
            Command1_Click
          Else: ClearScr: Exit Function
          End If
End Function
Public Function Collide()
On Error Resume Next
wallColide
For i = 0 To Shape1.Count - 1
    If Movement = i Then GoTo err
    If Shape1(Movement).Left = Shape1(i).Left Then
        If Shape1(Movement).Top = Shape1(i).Top Then
          If ShownEnd = True Then GoTo aa
          Timer1.Enabled = False
          Timer4.Enabled = False
          ShownEnd = True
          MsgBox "Game Over!" & vbCrLf & "---------------" & vbCrLf & "You Scored " & Label3.Caption & "!!!", vbOKOnly, "VB Snake"
          
          If HighestScore < Label3.Caption Then
            MsgBox "Guess What! You beat " & HighestPlayer & " by " & Label3.Caption & "/" & HighestScore
            Me.HighestScore = Label3.Caption
            Me.HighestPlayer = Me.Player
            SaveSettings
            Label4.Caption = "Current Highest: " & HighestPlayer & " with " & HighestScore
          Else
            If Int(Label3.Caption) > (HighestScore - 10) Then
                MsgBox "You were so close to beating him/her!, please try again!", vbOKOnly, "VB-Snake"
            Else
                MsgBox "Are you going to let " & HighestPlayer & " get away with winning all the time?" & vbCrLf & "if not, try to beat his/her score of " & HighestScore & "!!!", vbOKOnly, "VB-Snake"
            End If
            Command4.Visible = False
          End If
aa:
          ResetGame
          xxx = MsgBox("Play again?", vbYesNo, "VB-Snake")
          If xxx = vbYes Then
            Command1_Click
          Else: ClearScr: Exit Function
          End If
        End If
    End If
err:
Next i
End Function
Public Function GetMaxHeight() As Integer
Dim x As Integer
Do While x < Picture1.Height
DoEvents
x = x + 120
Loop
GetMaxHeight = 4200 'x
End Function
Public Function GetMaxWidth() As Integer
Dim x As Integer
Do While x < Picture1.Width
DoEvents
x = x + 120
Loop
GetMaxWidth = 5520 'x
End Function
Private Sub UpDown1_Change()
Text1.Text = UpDown1.Value
Timer1.Interval = (10 - Text1.Text) * 100
'Me.Caption = Timer1.Interval
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
If Label1.Caption = "1" Then Label1.Caption = "GO!": Timer1.Enabled = True: Timer3.Enabled = True: GoTo err
If Label1.Caption = "" Then Label1.Caption = "3"
Label1.Caption = Label1.Caption - 1
err:
End Sub

Private Sub Timer3_Timer()
NewGoal: NewGoal
Label1.Caption = ""
Timer3.Enabled = False
End Sub

Private Sub Timer4_Timer()
If Shape2.Left < 0 Then NewGoal
If Shape2.Top < 0 Then NewGoal
If Shape2.Top >= GetMaxHeight Then NewGoal
If Shape2.Left >= GetMaxWidth Then NewGoal
End Sub

Private Sub Timer5_Timer()

End Sub

