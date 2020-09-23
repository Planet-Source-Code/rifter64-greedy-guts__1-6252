VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FF0000&
   Caption         =   "GREEDY GUTS"
   ClientHeight    =   6225
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8910
   LinkTopic       =   "Form1"
   ScaleHeight     =   6225
   ScaleWidth      =   8910
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   120
      ScaleHeight     =   2025
      ScaleWidth      =   6945
      TabIndex        =   17
      Top             =   2760
      Width           =   6975
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Index           =   34
         X1              =   120
         X2              =   120
         Y1              =   720
         Y2              =   240
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Index           =   33
         X1              =   120
         X2              =   120
         Y1              =   1800
         Y2              =   1320
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Index           =   32
         X1              =   3360
         X2              =   3360
         Y1              =   720
         Y2              =   240
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Index           =   31
         X1              =   3600
         X2              =   3600
         Y1              =   720
         Y2              =   240
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Index           =   30
         X1              =   3600
         X2              =   3600
         Y1              =   1800
         Y2              =   1320
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Index           =   29
         X1              =   3360
         X2              =   3360
         Y1              =   1800
         Y2              =   1320
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Index           =   28
         X1              =   6840
         X2              =   6840
         Y1              =   720
         Y2              =   240
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Index           =   27
         X1              =   6840
         X2              =   6840
         Y1              =   1800
         Y2              =   1320
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Index           =   26
         X1              =   120
         X2              =   3360
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Index           =   25
         X1              =   120
         X2              =   3360
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Index           =   24
         X1              =   120
         X2              =   3360
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Index           =   23
         X1              =   3600
         X2              =   6840
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Index           =   22
         X1              =   3600
         X2              =   6840
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Index           =   21
         X1              =   3600
         X2              =   6840
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Index           =   20
         X1              =   3600
         X2              =   6840
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Index           =   4
         X1              =   120
         X2              =   3360
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Label lblB 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "Label4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   3600
         MouseIcon       =   "Trivia.frx":0000
         TabIndex        =   21
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label lblD 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   3600
         MouseIcon       =   "Trivia.frx":0152
         TabIndex        =   20
         Top             =   1320
         Width           =   3255
      End
      Begin VB.Label lblC 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   120
         MouseIcon       =   "Trivia.frx":02A4
         TabIndex        =   19
         Top             =   1320
         Width           =   3255
      End
      Begin VB.Label lblA 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   120
         MouseIcon       =   "Trivia.frx":03F6
         TabIndex        =   18
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Height          =   495
         Index           =   3
         Left            =   3600
         TabIndex        =   28
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Height          =   495
         Index           =   2
         Left            =   120
         TabIndex        =   27
         Top             =   1320
         Width           =   3255
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Height          =   495
         Index           =   1
         Left            =   3600
         TabIndex        =   26
         Top             =   1320
         Width           =   3255
      End
   End
   Begin VB.Label Pool 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "INSTA PICK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   3120
      TabIndex        =   29
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label Phone 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "GUESS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   5040
      TabIndex        =   24
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Fifty 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF0000&
      Caption         =   "POLL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   6240
      TabIndex        =   23
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label lblmess 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   22
      Top             =   5760
      Width           =   8655
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   1560
      Picture         =   "Trivia.frx":0548
      Top             =   2040
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   840
      Picture         =   "Trivia.frx":098A
      Top             =   2040
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "Trivia.frx":0C94
      Top             =   2040
      Width           =   480
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808080&
      FillColor       =   &H00000040&
      FillStyle       =   0  'Solid
      Height          =   2055
      Left            =   240
      Top             =   2880
      Width           =   6975
   End
   Begin VB.Label lblGoing 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "GOING FOR $ 100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   120
      TabIndex        =   16
      Top             =   5160
      Width           =   7095
   End
   Begin VB.Label right15 
      BackColor       =   &H0000FF00&
      Caption         =   "$ 3,276,800"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   7440
      TabIndex        =   15
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label right14 
      BackColor       =   &H0000FF00&
      Caption         =   "$ 1,639,600"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   7440
      TabIndex        =   14
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label right13 
      BackColor       =   &H0000FF00&
      Caption         =   "$ 819,800"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   7440
      TabIndex        =   13
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label right12 
      BackColor       =   &H0000FF00&
      Caption         =   "$ 409,600"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   7440
      TabIndex        =   12
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label right11 
      BackColor       =   &H0000FF00&
      Caption         =   "$ 204,800"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   7440
      TabIndex        =   11
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label right10 
      BackColor       =   &H0000FF00&
      Caption         =   "$ 102,400"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   7440
      TabIndex        =   10
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label right9 
      BackColor       =   &H0000FF00&
      Caption         =   "$ 51,200"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   7440
      TabIndex        =   9
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label right8 
      BackColor       =   &H0000FF00&
      Caption         =   "$ 25,600"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   7440
      TabIndex        =   8
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label right7 
      BackColor       =   &H0000FF00&
      Caption         =   "$ 12,800"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   7440
      TabIndex        =   7
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label right6 
      BackColor       =   &H0000FF00&
      Caption         =   "$ 6,400"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   7440
      TabIndex        =   6
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label right5 
      BackColor       =   &H0000FF00&
      Caption         =   "$ 3,200"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   7440
      TabIndex        =   5
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label right4 
      BackColor       =   &H0000FF00&
      Caption         =   "$ 1,600"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   7440
      TabIndex        =   4
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label right3 
      BackColor       =   &H0000FF00&
      Caption         =   "$ 800"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   7440
      TabIndex        =   3
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label right2 
      BackColor       =   &H0000FF00&
      Caption         =   "$ 400"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   7440
      TabIndex        =   2
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Label right1 
      BackColor       =   &H0000FF00&
      Caption         =   "$ 200"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   7440
      TabIndex        =   1
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   3
      X1              =   7320
      X2              =   8760
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   2
      X1              =   8760
      X2              =   7320
      Y1              =   5640
      Y2              =   5640
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   1
      X1              =   8760
      X2              =   8760
      Y1              =   120
      Y2              =   5640
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   0
      X1              =   7320
      X2              =   7320
      Y1              =   120
      Y2              =   5640
   End
   Begin VB.Label lblQuest 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6975
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00000040&
      BorderColor     =   &H00000040&
      FillColor       =   &H00000040&
      FillStyle       =   0  'Solid
      Height          =   1695
      Left            =   240
      Top             =   240
      Width           =   6975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Program made by RYAN HOLLETT FEB 28,2000
'You can use thuis source code to help you make your own trivias game
'but don't take this program and add questions and say that it is yours
'be creative and come up with your own game.
'The program ay have some spelling mistakes but who cares this is just a demo
Private Sub Check()
    Do While QuestionNum = OldQuest1 Or QuestionNum = OldQuest2 Or QuestionNum = OldQuest3 _
      Or QuestionNum = OldQuest4 Or QuestionNum = OldQuest5 Or QuestionNum = OldQuest6 _
      Or QuestionNum = OldQuest7 Or QuestionNum = OldQuest8 Or QuestionNum = OldQuest9 _
      Or QuestionNum = OldQuest10 Or QuestionNum = OldQuest11 Or QuestionNum = OldQuest12 _
      Or QuestionNum = OldQuest13 Or QuestionNum = OldQuest14
        QuestionNum = Int(Rnd * 38) + 1
    Loop
End Sub
Private Sub Swap()
    If Starts = 1 Then
        Let OldQuest1 = QuestionNum
    ElseIf Starts = 2 Then
        Let OldQuest2 = QuestionNum
    ElseIf Starts = 3 Then
        Let OldQuest3 = QuestionNum
    ElseIf Starts = 4 Then
        Let OldQuest4 = QuestionNum
    ElseIf Starts = 5 Then
        Let OldQuest5 = QuestionNum
    ElseIf Starts = 6 Then
        Let OldQuest6 = QuestionNum
    ElseIf Starts = 7 Then
        Let OldQuest7 = QuestionNum
    ElseIf Starts = 8 Then
        Let OldQuest8 = QuestionNum
    ElseIf Starts = 9 Then
        Let OldQuest9 = QuestionNum
    ElseIf Starts = 10 Then
        Let OldQuest10 = QuestionNum
    ElseIf Starts = 11 Then
        Let OldQuest11 = QuestionNum
    ElseIf Starts = 12 Then
        Let OldQuest12 = QuestionNum
    ElseIf Starts = 13 Then
        Let OldQuest13 = QuestionNum
    ElseIf Starts = 14 Then
        Let OldQuest14 = QuestionNum
    End If
End Sub
Private Sub Start()
    Let lblA.Visible = True
    Let lblB.Visible = True
    Let lblC.Visible = True
    Let lblD.Visible = True
    Randomize Timer
    QuestionNum = Int(Rnd * 38) + 1
    Starts = Starts + 1
    Call Swap
    Call Check
    If QuestionNum = 1 Then
        Call Quest1
    ElseIf QuestionNum = 2 Then
        Call Quest2
    ElseIf QuestionNum = 3 Then
        Call Quest3
    ElseIf QuestionNum = 4 Then
        Call Quest4
    ElseIf QuestionNum = 5 Then
        Call Quest5
    ElseIf QuestionNum = 6 Then
        Call Quest6
    ElseIf QuestionNum = 7 Then
        Call Quest7
    ElseIf QuestionNum = 8 Then
        Call Quest8
    ElseIf QuestionNum = 9 Then
        Call Quest9
    ElseIf QuestionNum = 10 Then
        Call Quest10
    ElseIf QuestionNum = 11 Then
        Call Quest11
    ElseIf QuestionNum = 12 Then
        Call Quest12
    ElseIf QuestionNum = 13 Then
        Call Quest13
    ElseIf QuestionNum = 14 Then
        Call Quest14
    ElseIf QuestionNum = 15 Then
        Call Quest15
    ElseIf QuestionNum = 16 Then
        Call Quest16
    ElseIf QuestionNum = 17 Then
        Call Quest17
    ElseIf QuestionNum = 18 Then
        Call Quest18
    ElseIf QuestionNum = 19 Then
        Call Quest19
    ElseIf QuestionNum = 20 Then
        Call Quest20
    ElseIf QuestionNum = 21 Then
        Call Quest21
    ElseIf QuestionNum = 22 Then
        Call Quest22
    ElseIf QuestionNum = 23 Then
        Call Quest23
    ElseIf QuestionNum = 24 Then
        Call Quest24
    ElseIf QuestionNum = 25 Then
        Call Quest25
    ElseIf QuestionNum = 26 Then
        Call Quest26
    ElseIf QuestionNum = 27 Then
        Call Quest27
    ElseIf QuestionNum = 28 Then
        Call Quest28
    ElseIf QuestionNum = 29 Then
        Call Quest29
    ElseIf QuestionNum = 30 Then
        Call Quest30
    ElseIf QuestionNum = 31 Then
        Call Quest31
    ElseIf QuestionNum = 32 Then
        Call Quest32
    ElseIf QuestionNum = 33 Then
        Call Quest33
    ElseIf QuestionNum = 34 Then
        Call Quest34
    ElseIf QuestionNum = 35 Then
        Call Quest35
    ElseIf QuestionNum = 36 Then
        Call Quest36
    ElseIf QuestionNum = 37 Then
        Call Quest37
    ElseIf QuestionNum = 38 Then
        Call Quest38
    End If
End Sub
Private Sub Fifty_Click()
    Form2.Show 1
    Let Fifty.Enabled = False
End Sub
Private Sub Fifty_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Let lblmess.Caption = "GENERATE A COMPUTER POLL"
End Sub
Private Sub Form_Load()
    Let Image3.Enabled = False
    Call Start
End Sub
Private Sub Quest1()
    Let lblQuest.Caption = "What team has won the most cups in the NHL since the begining of the stanley cup in 1893 ?"
    Let lblA.Caption = "A: MONTREAL"
    Let lblB.Caption = "B: TORONTO"
    Let lblC.Caption = "C: BOSTON"
    Let lblD.Caption = "D: NEY YORK RANGERS"
    Answer = "A"
End Sub
Private Sub Quest2()
    Let lblQuest.Caption = "What player has accumalated more assist in his career than any player had points in the NHL?"
    Let lblA.Caption = "A: BOBBY ORR"
    Let lblB.Caption = "B: WAYNE GRETZKY"
    Let lblC.Caption = "C: GORDIE HOWE"
    Let lblD.Caption = "D: MARIO LEMIEUX"
    Answer = "B"
End Sub
Private Sub Quest3()
    Let lblQuest.Caption = "Who has scored the most points as a european, right winger in the NHL ?"
    Let lblA.Caption = "A: JAMIR JAGR"
    Let lblB.Caption = "B: PETER STASTNY"
    Let lblC.Caption = "C: PAVEL BURE"
    Let lblD.Caption = "D: ALEXANDER MOGILNY"
    Answer = "A"
End Sub
Private Sub Quest4()
    Let lblQuest.Caption = "Who wore the first goaltending helmet in the NHL ?"
    Let lblA.Caption = "A: KEN DRYDEN"
    Let lblB.Caption = "B: PATRICK ROY"
    Let lblC.Caption = "C: EDDIE BOWER"
    Let lblD.Caption = "D: JAQUES PLANTE"
    Answer = "D"
End Sub
Private Sub Quest5()
    Let lblQuest.Caption = "What was the first year of the NHL ?"
    Let lblA.Caption = "A: 1923"
    Let lblB.Caption = "B: 1913"
    Let lblC.Caption = "C: 1917"
    Let lblD.Caption = "D: 1921"
    Answer = "C"
End Sub
Private Sub Quest6()
    Let lblQuest.Caption = "What was the fist name of Ozzy Osbourne band, BLACK SABBATH, first when they started in the music business ?"
    Let lblA.Caption = "A: OZZY AND THE ROCKERS"
    Let lblB.Caption = "B: THE BOOMERS"
    Let lblC.Caption = "C: EARTH RIDERS"
    Let lblD.Caption = "D: EARTH"
    Answer = "D"
End Sub
Private Sub Quest7()
    Let lblQuest.Caption = "In what year did the band METALLICA release their first album ?"
    Let lblA.Caption = "A: 1979"
    Let lblB.Caption = "B: 1982"
    Let lblC.Caption = "C: 1983"
    Let lblD.Caption = "D: 1980"
    Answer = "C"
End Sub
Private Sub Quest8()
    Let lblQuest.Caption = "What female singer has had more singles than the BEATLES ?"
    Let lblA.Caption = "A: WHITNEY HOUSTON"
    Let lblB.Caption = "B: MARIAH CAREY"
    Let lblC.Caption = "C: FAITH EVANS"
    Let lblD.Caption = "D: SHANIA TWAIN"
    Answer = "B"
End Sub
Private Sub Quest9()
    Let lblQuest.Caption = "What animal is the symbol of payboy ?"
    Let lblA.Caption = "A: RABBIT"
    Let lblB.Caption = "B: DOG"
    Let lblC.Caption = "C: TEDDY BEAR"
    Let lblD.Caption = "D: CAT"
    Answer = "A"
End Sub
Private Sub Quest10()
    Let lblQuest.Caption = "What currency is used in CHINA ?"
    Let lblA.Caption = "A: ROOPLE"
    Let lblB.Caption = "B: YUAN"
    Let lblC.Caption = "C: YEN"
    Let lblD.Caption = "D: FRANK"
    Answer = "B"
End Sub
Private Sub Quest11()
    Let lblQuest.Caption = "What france owned island is off of the coast of NEWFOUNDLAND ?"
    Let lblA.Caption = "A: ST. PIERRE"
    Let lblB.Caption = "B: FRANCOUIX"
    Let lblC.Caption = "C: TROUIS RIVERES"
    Let lblD.Caption = "D: GRANDE ISLANDE"
    Answer = "A"
End Sub
Private Sub Quest12()
    Let lblQuest.Caption = "Who was the first NHL player to make the curved blade ?"
    Let lblA.Caption = "A: STAN MIKITA"
    Let lblB.Caption = "B: BOBBY ORR"
    Let lblC.Caption = "C: BOBBY HULL"
    Let lblD.Caption = "D: ANDY BATHGATE"
    Answer = "D"
End Sub
Private Sub Quest13()
    Let lblQuest.Caption = "What is more ?"
    Let lblA.Caption = "A: 1 Kb"
    Let lblB.Caption = "B: 1 MB"
    Let lblC.Caption = "C: 3 MB"
    Let lblD.Caption = "D: 230 Mb"
    Answer = "D"
End Sub
Private Sub Quest14()
    Let lblQuest.Caption = "How many spots does ten dalmations have ?"
    Let lblA.Caption = "A: 1010"
    Let lblB.Caption = "B: 101"
    Let lblC.Caption = "C: 110"
    Let lblD.Caption = "D: 10101"
    Answer = "A"
End Sub
Private Sub Quest15()
    Let lblQuest.Caption = "What sport uses a TEE, ROCK and CRAMPIT  ?"
    Let lblA.Caption = "A: GOLF"
    Let lblB.Caption = "B: BOXING"
    Let lblC.Caption = "C: CURLING"
    Let lblD.Caption = "D: HOCKEY"
    Answer = "C"
End Sub
Private Sub Quest16()
    Let lblQuest.Caption = "Who was the creator of E.T. ?"
    Let lblA.Caption = "A: STEVEN SPEILBERG"
    Let lblB.Caption = "B: STEVEN KING"
    Let lblC.Caption = "C: GEORGE LUCAS"
    Let lblD.Caption = "D: JAMES CAMERON"
    Answer = "A"
End Sub
Private Sub Quest17()
    Let lblQuest.Caption = "Who said 'I CAN'T TELL A LIE' ?"
    Let lblA.Caption = "A: STEVEN SPEILBERG"
    Let lblB.Caption = "B: GEORGE WASHINGTON"
    Let lblC.Caption = "C: BILL COSBY"
    Let lblD.Caption = "D: JOHN F. KENNEDY"
    Answer = "B"
End Sub
Private Sub Quest18()
    Let lblQuest.Caption = "What four standard standard playing cards are usually red ?"
    Let lblA.Caption = "A: HEARTS, SPADES"
    Let lblB.Caption = "B: SPADES, DIAMONDS"
    Let lblC.Caption = "C: HEARTS, DIAMONDS"
    Let lblD.Caption = "D: SPADES, CLUBS"
    Answer = "C"
End Sub
Private Sub Quest19()
    Let lblQuest.Caption = "What number is directly 180 degrees from 11 O'clock ?"
    Let lblA.Caption = "A: 6"
    Let lblB.Caption = "B: 3"
    Let lblC.Caption = "C: 4"
    Let lblD.Caption = "D: 5"
    Answer = "D"
End Sub
Private Sub Quest20()
    Let lblQuest.Caption = "The phrase 'The quick brown fox hops over the lazy frog' is an example of what ?"
    Let lblA.Caption = "A: SYNONYM"
    Let lblB.Caption = "B: ALL LETTERS OF ALPHABET"
    Let lblC.Caption = "C: LYMRIC"
    Let lblD.Caption = "D: RHYME"
    Answer = "B"
End Sub
Private Sub Quest21()
    Let lblQuest.Caption = "What musical instrument is known as 'The Licorice Stick' ?"
    Let lblA.Caption = "A: TRUMPET"
    Let lblB.Caption = "B: CLARINET"
    Let lblC.Caption = "C: FLOOTE"
    Let lblD.Caption = "D: SAXOPHONE"
    Answer = "B"
End Sub
Private Sub Quest22()
    Let lblQuest.Caption = "What actor starred as T.V.'s detective Dan August ?"
    Let lblA.Caption = "A: ROBERT ULRICH"
    Let lblB.Caption = "B: BURT RENYOLDS"
    Let lblC.Caption = "C: MORGAN FREEMAN"
    Let lblD.Caption = "D: DAVID SPADE"
    Answer = "B"
End Sub
Private Sub Quest24()
    Let lblQuest.Caption = "Which era were dinosaurs alive ?"
    Let lblA.Caption = "A: MESOZOIC"
    Let lblB.Caption = "B: CENOZOIC"
    Let lblC.Caption = "C: PROTEROZOIC"
    Let lblD.Caption = "D: ARCHEOZOIC"
    Answer = "A"
End Sub
Private Sub Quest25()
    Let lblQuest.Caption = "Who normally live in monostaries ?"
    Let lblA.Caption = "A: MONGOLS"
    Let lblB.Caption = "B: MONSTERS"
    Let lblC.Caption = "C: MONEYLENDERS"
    Let lblD.Caption = "D: MONKS"
    Answer = "D"
End Sub
Private Sub Quest26()
    Let lblQuest.Caption = "What type of animal is the Cheetos mascot ?"
    Let lblA.Caption = "A: CHICKEN"
    Let lblB.Caption = "B: FROG"
    Let lblC.Caption = "C: CHEETAH"
    Let lblD.Caption = "D: BEAR"
    Answer = "C"
End Sub
Private Sub Quest27()
    Let lblQuest.Caption = "Gingivitis cause inflammation in what part of the body ?"
    Let lblA.Caption = "A: LIPS"
    Let lblB.Caption = "B: ARM"
    Let lblC.Caption = "C: EAR"
    Let lblD.Caption = "D: GUMS"
    Answer = "D"
End Sub
Private Sub Quest28()
    Let lblQuest.Caption = "What did the 12 people in the 1967 movie 'The Dirty Dozen' have in common ?"
    Let lblA.Caption = "A: COUSINS"
    Let lblB.Caption = "B: BROTHERS"
    Let lblC.Caption = "C: BANK ROBBERS"
    Let lblD.Caption = "D: CONVICTS"
    Answer = "D"
End Sub
Private Sub Quest29()
    Let lblQuest.Caption = "What town did CLARK KENT grow up in ?"
    Let lblA.Caption = "A: METROPOLIS"
    Let lblB.Caption = "B: SMALLVILLE"
    Let lblC.Caption = "C: COUNTRYSIDE"
    Let lblD.Caption = "D: OUTBACK"
    Answer = "B"
End Sub
Private Sub Quest30()
    Let lblQuest.Caption = "Where does homer simpson work ?"
    Let lblA.Caption = "A: BAR"
    Let lblB.Caption = "B: MALL"
    Let lblC.Caption = "C: NUCLEAR POWER PLANT"
    Let lblD.Caption = "D: QUICKY MART"
    Answer = "C"
End Sub
Private Sub Quest31()
    Let lblQuest.Caption = "What puncuation mark is used to show emission of or more letters in a contraction ?"
    Let lblA.Caption = "A: OPEN PARENTHESIS"
    Let lblB.Caption = "B: SEMI-COLON"
    Let lblC.Caption = "C: APOSTROPHE"
    Let lblD.Caption = "D: EXCLAMATION POINT"
    Answer = "C"
End Sub
Private Sub Quest32()
    Let lblQuest.Caption = "Masonry is primarily built with this ?"
    Let lblA.Caption = "A: BRICK"
    Let lblB.Caption = "B: CLAY"
    Let lblC.Caption = "C: WOOD"
    Let lblD.Caption = "D: STONE"
    Answer = "D"
End Sub
Private Sub Quest33()
    Let lblQuest.Caption = "Which PENUTS character was proud or their curly hair?"
    Let lblA.Caption = "A: MARCIE"
    Let lblB.Caption = "B: CHARLIE BROWN"
    Let lblC.Caption = "C: FREIDA"
    Let lblD.Caption = "D: LINUS"
    Answer = "C"
End Sub
Private Sub Quest23()
    Let lblQuest.Caption = "What university did TIGER WOODS attend before becoming pro ?"
    Let lblA.Caption = "A: YALE"
    Let lblB.Caption = "B: UCLA"
    Let lblC.Caption = "C: MIAMI STATE "
    Let lblD.Caption = "D: STANFORD"
    Answer = "D"
End Sub
Private Sub Quest34()
    Let lblQuest.Caption = "What major city calls itself 'The Windy City' ?"
    Let lblA.Caption = "A: DETROIT"
    Let lblB.Caption = "B: SEATTLE"
    Let lblC.Caption = "C: CHEYENE"
    Let lblD.Caption = "D: CHICAGO"
    Answer = "D"
End Sub
Private Sub Quest35()
    Let lblQuest.Caption = "What is dolomite ?"
    Let lblA.Caption = "A: MINERAL FOUND BY LIMESTONE"
    Let lblB.Caption = "B: SOIL"
    Let lblC.Caption = "C: ROCK AROUND CORAL"
    Let lblD.Caption = "D: SCUM ON POND"
    Answer = "A"
End Sub
Private Sub Quest36()
    Let lblQuest.Caption = "What are the names of Angelica's parents in the Rugrats ?"
    Let lblA.Caption = "A: STEW AND DEEDEE"
    Let lblB.Caption = "B: DREW AND CHARLOTTE"
    Let lblC.Caption = "C: TOMMY AND SUZY"
    Let lblD.Caption = "D: GEOFF AND MILDRED"
    Answer = "B"
End Sub
Private Sub Quest37()
    Let lblQuest.Caption = "In the bible who's name was changed to Isreal ?"
    Let lblA.Caption = "A: JOSHUA"
    Let lblB.Caption = "B: NOAH"
    Let lblC.Caption = "C: MOSES"
    Let lblD.Caption = "D: JACOB"
    Answer = "D"
End Sub
Private Sub Quest38()
    Let lblQuest.Caption = "What is the last name of the beardless Z.Z.Top band member ?"
    Let lblA.Caption = "A: ZEZEEL"
    Let lblB.Caption = "B: TOPHAM"
    Let lblC.Caption = "C: BEARD"
    Let lblD.Caption = "D: ZACHARIAS"
    Answer = "C"
End Sub
Private Sub Image1_Click()
        Let Phone.Enabled = True
        Let Fifty.Enabled = True
        Let Pool.Enabled = True
        Let lblA.Enabled = False
        Let lblB.Enabled = False
        Let lblC.Enabled = False
        Let lblD.Enabled = False
        Let Image1.Enabled = False
        Let Image3.Enabled = True
        Let lblA.Caption = ""
        Let lblB.Caption = ""
        Let lblC.Caption = ""
        Let lblD.Caption = ""
        Let lblQuest.Font.Size = 24
        Let lblQuest.BackColor = vbYellow
        If RightAns > 0 Then
            Let lblQuest.Caption = "YOU HAVE DECIDED TO TAKE THE MONEY"
            lblGoing = "YOUR TOTAL PRIZE IS " & TotPrize
        Else
            lblGoing.Caption = "COME ON IT'S NOT THAT HARD"
            Let lblQuest.Caption = "WHY DID YOU CHICKEN OUT ?"
        End If
End Sub
Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Let lblmess.Caption = "KEEP THE WINNINGS AND QUIT"
End Sub
Private Sub Image2_Click()
    End
End Sub
Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Let lblmess.Caption = "EXIT THE PROGRAM AND RECIEVE NO WINNINGS"
End Sub
Private Sub Image3_Click()
    Let Phone.Enabled = True
    Let Fifty.Enabled = True
    Let Pool.Enabled = True
    Let RightAns = 0
    Let Image1.Enabled = True
    Let lblA.Enabled = True
    Let lblB.Enabled = True
    Let lblC.Enabled = True
    Let lblD.Enabled = True
    Let lblA.Caption = ""
    Let lblB.Caption = ""
    Let lblC.Caption = ""
    Let lblD.Caption = ""
    Let lblQuest.Font.Size = 12
    Let lblQuest.BackColor = vbGreen
    Let Image1.Enabled = True
    Let Wrongs = 0
    Call Start
End Sub
Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Let lblmess.Caption = "NEW GAME"
End Sub

Private Sub Phone_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Let lblmess.Caption = "TAKE A GUESS ON THE ANSWER"
End Sub
Private Sub pool_Click()
    Dim Insta As Integer
    Insta = Int(Rnd * 4) + 1
    If Insta = 1 Then
        If Answer = "A" Then
            lblA.BackColor = vbYellow
            Call Win
        Else
            MsgBox "INSTA PICK PICKED 'A'"
            Let lblA.Visible = False
        End If
    ElseIf Insta = 2 Then
        If Answer = "B" Then
            lblB.BackColor = vbYellow
            Call Win
        Else
            MsgBox "INSTA PICK PICKED 'B'"
            Let lblB.Visible = False
        End If
    ElseIf Insta = 3 Then
        If Answer = "C" Then
            lblC.BackColor = vbYellow
            Call Win
        Else
            MsgBox "INSTA PICK PICKED 'C'"
            Let lblC.Visible = False
        End If
    ElseIf Insta = 4 Then
        If Answer = "D" Then
            lblD.BackColor = vbYellow
            Call Win
        Else
            MsgBox "INSTA PICK PICKED 'D'"
            Let lblD.Visible = False
        End If
    End If
    Let Pool.Enabled = False
End Sub
Private Sub lblA_Click()
    LabelAnswer = "A"
    If LabelAnswer = Answer Then
        Call Win
    Else
        Call Lose
    End If
End Sub
Private Sub lblb_Click()
    LabelAnswer = "B"
    If LabelAnswer = Answer Then
        Call Win
    Else
        Call Lose
    End If
End Sub
Private Sub lblC_Click()
    LabelAnswer = "C"
    If LabelAnswer = Answer Then
        Call Win
    Else
        Call Lose
    End If
End Sub
Private Sub lblD_Click()
    LabelAnswer = "D"
    If LabelAnswer = Answer Then
        Call Win
    Else
        Call Lose
    End If
End Sub
Public Sub Win()
    MsgBox "YOU'VE GOT THE RIGHT ANSWER"
    Let RightAns = RightAns + 1
    Call Swap
    Call Start
    Call WinningBar
End Sub
Private Sub WinningBar()
    If RightAns = 1 Then
        right1.BackColor = vbBlue
        right2.BackColor = vbGreen
        right3.BackColor = vbGreen
        right4.BackColor = vbGreen
        right5.BackColor = vbGreen
        right6.BackColor = vbGreen
        right7.BackColor = vbGreen
        right8.BackColor = vbGreen
        right9.BackColor = vbGreen
        right10.BackColor = vbGreen
        right11.BackColor = vbGreen
        right12.BackColor = vbGreen
        right13.BackColor = vbGreen
        right14.BackColor = vbGreen
        right15.BackColor = vbGreen
        lblGoing = "GOING FOR " & right2.Caption
        TotPrize = right1.Caption
        Let OldQuest1 = QuestionNum
    ElseIf RightAns = 2 Then
        right1.BackColor = vbGreen
        right2.BackColor = vbBlue
        right3.BackColor = vbGreen
        right4.BackColor = vbGreen
        right5.BackColor = vbGreen
        right6.BackColor = vbGreen
        right7.BackColor = vbGreen
        right8.BackColor = vbGreen
        right9.BackColor = vbGreen
        right10.BackColor = vbGreen
        right11.BackColor = vbGreen
        right12.BackColor = vbGreen
        right13.BackColor = vbGreen
        right14.BackColor = vbGreen
        right15.BackColor = vbGreen
        lblGoing = "GOING FOR " & right3.Caption
        TotPrize = right2.Caption
        Let OldQuest2 = QuestionNum
    ElseIf RightAns = 3 Then
        right1.BackColor = vbGreen
        right2.BackColor = vbGreen
        right3.BackColor = vbBlue
        right4.BackColor = vbGreen
        right5.BackColor = vbGreen
        right6.BackColor = vbGreen
        right7.BackColor = vbGreen
        right8.BackColor = vbGreen
        right9.BackColor = vbGreen
        right10.BackColor = vbGreen
        right11.BackColor = vbGreen
        right12.BackColor = vbGreen
        right13.BackColor = vbGreen
        right14.BackColor = vbGreen
        right15.BackColor = vbGreen
        lblGoing = "GOING FOR " & right4.Caption
        TotPrize = right3.Caption
        Let OldQuest3 = QuestionNum
    ElseIf RightAns = 4 Then
        right1.BackColor = vbGreen
        right2.BackColor = vbGreen
        right3.BackColor = vbGreen
        right4.BackColor = vbBlue
        right5.BackColor = vbGreen
        right6.BackColor = vbGreen
        right7.BackColor = vbGreen
        right8.BackColor = vbGreen
        right9.BackColor = vbGreen
        right10.BackColor = vbGreen
        right11.BackColor = vbGreen
        right12.BackColor = vbGreen
        right13.BackColor = vbGreen
        right14.BackColor = vbGreen
        right15.BackColor = vbGreen
        lblGoing = "GOING FOR " & right5.Caption
        TotPrize = right4.Caption
        Let OldQuest4 = QuestionNum
    ElseIf RightAns = 5 Then
        right1.BackColor = vbGreen
        right2.BackColor = vbGreen
        right3.BackColor = vbGreen
        right4.BackColor = vbGreen
        right5.BackColor = vbBlue
        right6.BackColor = vbGreen
        right7.BackColor = vbGreen
        right8.BackColor = vbGreen
        right9.BackColor = vbGreen
        right10.BackColor = vbGreen
        right11.BackColor = vbGreen
        right12.BackColor = vbGreen
        right13.BackColor = vbGreen
        right14.BackColor = vbGreen
        right15.BackColor = vbGreen
        lblGoing = "GOING FOR " & right6.Caption
        TotPrize = right5.Caption
        Let OldQuest5 = QuestionNum
    ElseIf RightAns = 6 Then
        right1.BackColor = vbGreen
        right2.BackColor = vbGreen
        right3.BackColor = vbGreen
        right4.BackColor = vbGreen
        right5.BackColor = vbGreen
        right6.BackColor = vbBlue
        right7.BackColor = vbGreen
        right8.BackColor = vbGreen
        right9.BackColor = vbGreen
        right10.BackColor = vbGreen
        right11.BackColor = vbGreen
        right12.BackColor = vbGreen
        right13.BackColor = vbGreen
        right14.BackColor = vbGreen
        right15.BackColor = vbGreen
        lblGoing = "GOING FOR " & right7.Caption
        TotPrize = right6.Caption
        Let OldQuest6 = QuestionNum
    ElseIf RightAns = 7 Then
        right1.BackColor = vbGreen
        right2.BackColor = vbGreen
        right3.BackColor = vbGreen
        right4.BackColor = vbGreen
        right5.BackColor = vbGreen
        right6.BackColor = vbGreen
        right7.BackColor = vbBlue
        right8.BackColor = vbGreen
        right9.BackColor = vbGreen
        right10.BackColor = vbGreen
        right11.BackColor = vbGreen
        right12.BackColor = vbGreen
        right13.BackColor = vbGreen
        right14.BackColor = vbGreen
        right15.BackColor = vbGreen
        lblGoing = "GOING FOR " & right8.Caption
        TotPrize = right7.Caption
        Let OldQuest7 = QuestionNum
    ElseIf RightAns = 8 Then
        right1.BackColor = vbGreen
        right2.BackColor = vbGreen
        right3.BackColor = vbGreen
        right4.BackColor = vbGreen
        right5.BackColor = vbGreen
        right6.BackColor = vbGreen
        right7.BackColor = vbGreen
        right8.BackColor = vbBlue
        right9.BackColor = vbGreen
        right10.BackColor = vbGreen
        right11.BackColor = vbGreen
        right12.BackColor = vbGreen
        right13.BackColor = vbGreen
        right14.BackColor = vbGreen
        right15.BackColor = vbGreen
        lblGoing = "GOING FOR " & right9.Caption
        TotPrize = right8.Caption
        Let OldQuest8 = QuestionNum
    ElseIf RightAns = 9 Then
        right1.BackColor = vbGreen
        right2.BackColor = vbGreen
        right3.BackColor = vbGreen
        right4.BackColor = vbGreen
        right5.BackColor = vbGreen
        right6.BackColor = vbGreen
        right7.BackColor = vbGreen
        right8.BackColor = vbGreen
        right9.BackColor = vbBlue
        right10.BackColor = vbGreen
        right11.BackColor = vbGreen
        right12.BackColor = vbGreen
        right13.BackColor = vbGreen
        right14.BackColor = vbGreen
        right15.BackColor = vbGreen
        lblGoing = "GOING FOR " & right10.Caption
        TotPrize = right9.Caption
        Let OldQuest9 = QuestionNum
    ElseIf RightAns = 10 Then
        right1.BackColor = vbGreen
        right2.BackColor = vbGreen
        right3.BackColor = vbGreen
        right4.BackColor = vbGreen
        right5.BackColor = vbGreen
        right6.BackColor = vbGreen
        right7.BackColor = vbGreen
        right8.BackColor = vbGreen
        right9.BackColor = vbGreen
        right10.BackColor = vbBlue
        right11.BackColor = vbGreen
        right12.BackColor = vbGreen
        right13.BackColor = vbGreen
        right14.BackColor = vbGreen
        right15.BackColor = vbGreen
        lblGoing = "GOING FOR " & right11.Caption
        TotPrize = right10.Caption
        Let OldQuest10 = QuestionNum
    ElseIf RightAns = 11 Then
        right1.BackColor = vbGreen
        right2.BackColor = vbGreen
        right3.BackColor = vbGreen
        right4.BackColor = vbGreen
        right5.BackColor = vbGreen
        right6.BackColor = vbGreen
        right7.BackColor = vbGreen
        right8.BackColor = vbGreen
        right9.BackColor = vbGreen
        right10.BackColor = vbGreen
        right11.BackColor = vbBlue
        right12.BackColor = vbGreen
        right13.BackColor = vbGreen
        right14.BackColor = vbGreen
        right15.BackColor = vbGreen
        lblGoing = "GOING FOR " & right12.Caption
        TotPrize = right11.Caption
        Let OldQuest11 = QuestionNum
    ElseIf RightAns = 12 Then
        right1.BackColor = vbGreen
        right2.BackColor = vbGreen
        right3.BackColor = vbGreen
        right4.BackColor = vbGreen
        right5.BackColor = vbGreen
        right6.BackColor = vbGreen
        right7.BackColor = vbGreen
        right8.BackColor = vbGreen
        right9.BackColor = vbGreen
        right10.BackColor = vbGreen
        right11.BackColor = vbGreen
        right12.BackColor = vbBlue
        right13.BackColor = vbGreen
        right14.BackColor = vbGreen
        right15.BackColor = vbGreen
        lblGoing = "GOING FOR " & right13.Caption
        TotPrize = right12.Caption
        Let OldQuest12 = QuestionNum
    ElseIf RightAns = 13 Then
        right1.BackColor = vbGreen
        right2.BackColor = vbGreen
        right3.BackColor = vbGreen
        right4.BackColor = vbGreen
        right5.BackColor = vbGreen
        right6.BackColor = vbGreen
        right7.BackColor = vbGreen
        right8.BackColor = vbGreen
        right9.BackColor = vbGreen
        right10.BackColor = vbGreen
        right11.BackColor = vbGreen
        right12.BackColor = vbGreen
        right13.BackColor = vbBlue
        right14.BackColor = vbGreen
        right15.BackColor = vbGreen
        lblGoing = "GOING FOR " & right14.Caption
        TotPrize = right13.Caption
        Let OldQuest13 = QuestionNum
    ElseIf RightAns = 14 Then
        right1.BackColor = vbGreen
        right2.BackColor = vbGreen
        right3.BackColor = vbGreen
        right4.BackColor = vbGreen
        right5.BackColor = vbGreen
        right6.BackColor = vbGreen
        right7.BackColor = vbGreen
        right8.BackColor = vbGreen
        right9.BackColor = vbGreen
        right10.BackColor = vbGreen
        right11.BackColor = vbGreen
        right12.BackColor = vbGreen
        right13.BackColor = vbGreen
        right14.BackColor = vbBlue
        right15.BackColor = vbGreen
        lblGoing = "GOING FOR " & right15.Caption
        TotPrize = right14.Caption
        Let OldQuest14 = QuestionNum
    ElseIf RightAns = 15 Then
        right1.BackColor = vbGreen
        right2.BackColor = vbGreen
        right3.BackColor = vbGreen
        right4.BackColor = vbGreen
        right5.BackColor = vbGreen
        right6.BackColor = vbGreen
        right7.BackColor = vbGreen
        right8.BackColor = vbGreen
        right9.BackColor = vbGreen
        right10.BackColor = vbGreen
        right11.BackColor = vbGreen
        right12.BackColor = vbGreen
        right13.BackColor = vbGreen
        right14.BackColor = vbGreen
        right15.BackColor = vbBlue
        Let lblA.Enabled = False
        Let lblB.Enabled = False
        Let lblC.Enabled = False
        Let lblD.Enabled = False
        Let lblA.Caption = ""
        Let lblB.Caption = ""
        Let lblC.Caption = ""
        Let lblD.Caption = ""
        Let lblQuest.Font.Size = 24
        Let lblQuest.BackColor = vbYellow
        Let lblQuest.Caption = "YOUR THE GRAND PRIZE WINNER"
        lblGoing = "YOU DONE IT! CONGRATS"
        TotPrize = right15.Caption
        Let Image1.Enabled = False
        Let Image3.Enabled = True
        Let OldQuest15 = QuestionNum
        Let Pool.Enabled = False
        Let Phone.Enabled = False
        Let Fifty.Enabled = False
    End If
End Sub
Private Sub Lose()
    MsgBox "WRONG ANSWER"
    Let Fifty.Enabled = False
    Let Pool.Enabled = False
    Let Phone.Enabled = False
    Let lblA.Enabled = False
    Let lblB.Enabled = False
    Let lblC.Enabled = False
    Let lblD.Enabled = False
    Let lblA.Visible = True
    Let lblB.Visible = True
    Let lblC.Visible = True
    Let lblD.Visible = True
    Let lblA.Caption = ""
    Let lblB.Caption = ""
    Let lblC.Caption = ""
    Let lblD.Caption = ""
    Let lblQuest.Font.Size = 24
    Let lblQuest.BackColor = vbYellow
    Let lblQuest.Caption = "SORRY TRY NEXT TIME"
    Let lblGoing.Caption = "YOU LOST AND YOU WON NO MONEY"
    Let Image3.Enabled = True
End Sub
Private Sub lblQuest_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Let lblA.BackColor = vbGreen
    Let lblB.BackColor = vbGreen
    Let lblC.BackColor = vbGreen
    Let lblD.BackColor = vbGreen
    Let lblmess.Caption = ""
End Sub
Private Sub lbla_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Let lblA.BackColor = vbYellow
    Let lblB.BackColor = vbGreen
    Let lblC.BackColor = vbGreen
    Let lblD.BackColor = vbGreen
    Let lblmess.Caption = ""
End Sub
Private Sub lblB_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Let lblA.BackColor = vbGreen
    Let lblB.BackColor = vbYellow
    Let lblC.BackColor = vbGreen
    Let lblD.BackColor = vbGreen
    Let lblmess.Caption = ""
End Sub
Private Sub lblC_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Let lblA.BackColor = vbGreen
    Let lblB.BackColor = vbGreen
    Let lblC.BackColor = vbYellow
    Let lblD.BackColor = vbGreen
    Let lblmess.Caption = ""
End Sub
Private Sub lblD_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Let lblA.BackColor = vbGreen
    Let lblB.BackColor = vbGreen
    Let lblC.BackColor = vbGreen
    Let lblD.BackColor = vbYellow
    Let lblmess.Caption = ""
End Sub
Private Sub form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Let lblA.BackColor = vbGreen
    Let lblB.BackColor = vbGreen
    Let lblC.BackColor = vbGreen
    Let lblD.BackColor = vbGreen
    Let lblmess.Caption = ""
End Sub
Private Sub Phone_Click()
    Form3.Show 1
    Phone.Enabled = False
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Let lblA.BackColor = vbGreen
    Let lblB.BackColor = vbGreen
    Let lblC.BackColor = vbGreen
    Let lblD.BackColor = vbGreen
    Let lblmess.Caption = ""
End Sub
Private Sub Pool_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Let lblmess.Caption = "AN INSTANT PICK"
End Sub
