VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "New Game"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3825
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   3825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      Caption         =   "Quick Info"
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   3480
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   3765
      TabIndex        =   11
      Top             =   0
      Width           =   3825
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "By Steve Bailey (30023070)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   600
         Width           =   2655
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "VB Snake"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   2655
      End
      Begin VB.Image Image1 
         Height          =   900
         Left            =   2040
         Picture         =   "Form2.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1680
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&OK"
      Height          =   375
      Left            =   1440
      TabIndex        =   9
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2640
      TabIndex        =   8
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "New Game (Options...)"
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   3615
      Begin VB.CheckBox Check1 
         Caption         =   "Crazy Colorz"
         Height          =   255
         Left            =   1080
         TabIndex        =   10
         ToolTipText     =   "Interchanges between colors during game. (verry nice effect!)"
         Top             =   1920
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.ListBox List1 
         Height          =   840
         ItemData        =   "Form2.frx":1A60
         Left            =   1080
         List            =   "Form2.frx":1A6A
         TabIndex        =   7
         Top             =   960
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1080
         TabIndex        =   5
         Text            =   "Joe Bloggs"
         Top             =   240
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   1
         Text            =   "9"
         Top             =   600
         Width           =   1215
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   255
         Left            =   2400
         TabIndex        =   2
         Top             =   600
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         _Version        =   327681
         Value           =   9
         Max             =   13
         Min             =   1
         Enabled         =   -1  'True
      End
      Begin VB.Label Label3 
         Caption         =   "Maze"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Your Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Speed"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   1095
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.Command4.Visible = False
Unload Me
End Sub

'NewGoal: NewGoal
Private Sub Command2_Click()
Form1.ClearScr
Form1.Label1.Caption = "3"
Form1.Label3.Caption = "0"

    If Int(Text1.Text) < 10 Then
    Form1.Timer1.Interval = (10 - Text1.Text) * 100
    ElseIf Text1.Text = "10" Then
    Form1.Timer1.Interval = 75
    ElseIf Text1.Text = "11" Then
    Form1.Timer1.Interval = 50
    ElseIf Text1.Text = "12" Then
    Form1.Timer1.Interval = 25
    ElseIf Text1.Text = "13" Then
    Form1.Timer1.Interval = 10
    End If
    If Text1.Text > 10 Then
        x = MsgBox("Are you sure you want to play at this speed? " & (Form1.Timer1.Interval) & " Moves A Second?", vbYesNo, "VB-Snake")
        If x = vbNo Then Exit Sub
    End If
Form1.Timer2.Enabled = True
Form1.SmallScore = Text1.Text
Form1.Command4.Visible = True
Form1.Player = Text2.Text
Form1.Check1.Value = Check1.Value

Unload_Walls
Select Case List1.ListIndex
Case 1: DoBox
'Case 2: doshrapnel
'Case 3: dospacyshrapnel
End Select

Unload Me
End Sub

Public Function Unload_Walls()
For i = 1 To Form1.wall.Count - 1
Unload Form1.wall(i)
Next i
Form1.wall(0).Visible = False
End Function

Public Function DoBox()
Dim i As Integer
'On Error Resume Next
With Form1
.wall(0).Left = 0
.wall(0).Top = 0
.wall(0).Visible = True
    For x = 120 To 5520 Step 120
        i = i + 1
        Load .wall(i)
        .wall(i).Top = 0
        .wall(i).Left = x
        .wall(i).Visible = True
    Next x
    For j = 120 To 4200 Step 120
        i = i + 1
        Load .wall(i)
        .wall(i).Top = j
        .wall(i).Left = 0
        .wall(i).Visible = True
    Next j
    For k = 120 To 5520 Step 120
        i = i + 1
        Load .wall(i)
        .wall(i).Top = 4200
        .wall(i).Left = k
        .wall(i).Visible = True
    Next k
    For l = 120 To 4200 Step 120
        i = i + 1
        Load .wall(i)
        .wall(i).Top = l
        .wall(i).Left = 5520
        .wall(i).Visible = True
    Next l
End With
End Function

Private Sub Command3_Click()
MsgBox "Remember, you use the Numpad to control you snake, and the more food u scoff" & vbCrLf & " the larger your snake gets. remember to not get tangled :P", vbOKOnly
End Sub

Private Sub Form_Load()
Form1.Show
End Sub

Private Sub UpDown1_Change()
Text1.Text = UpDown1.Value
End Sub
