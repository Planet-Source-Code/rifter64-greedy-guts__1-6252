VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00000000&
   Caption         =   "TAKE A GUESS"
   ClientHeight    =   2280
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4455
   LinkTopic       =   "Form3"
   ScaleHeight     =   2280
   ScaleWidth      =   4455
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "CLOSE"
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Take a guess, If you're wrong nothing happens, if you're right you win."
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
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   4215
   End
   Begin VB.Label Gd 
      BackColor       =   &H00000000&
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   3360
      TabIndex        =   4
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Gc 
      BackColor       =   &H00000000&
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      TabIndex        =   3
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Gb 
      BackColor       =   &H00000000&
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   1200
      TabIndex        =   2
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Ga 
      BackColor       =   &H00000000&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   975
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ga_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Let Ga.BackColor = vbYellow
    Let Gb.BackColor = vbBlack
    Let Gc.BackColor = vbBlack
    Let Gd.BackColor = vbBlack
End Sub
Private Sub gB_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Let Ga.BackColor = vbBlack
    Let Gb.BackColor = vbYellow
    Let Gc.BackColor = vbBlack
    Let Gd.BackColor = vbBlack
End Sub
Private Sub gC_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Let Ga.BackColor = vbBlack
    Let Gb.BackColor = vbBlack
    Let Gc.BackColor = vbYellow
    Let Gd.BackColor = vbBlack
End Sub
Private Sub Gd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Let Ga.BackColor = vbBlack
    Let Gb.BackColor = vbBlack
    Let Gc.BackColor = vbBlack
    Let Gd.BackColor = vbYellow
End Sub
Private Sub form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Let Ga.BackColor = vbBlack
    Let Gb.BackColor = vbBlack
    Let Gc.BackColor = vbBlack
    Let Gd.BackColor = vbBlack
End Sub
Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Let Ga.BackColor = vbBlack
    Let Gb.BackColor = vbBlack
    Let Gc.BackColor = vbBlack
    Let Gd.BackColor = vbBlack
End Sub
Private Sub gA_Click()
    LabelAnswer = "A"
    If LabelAnswer = Answer Then
        Call Form1.Win
    End If
    Unload Form3
End Sub
Private Sub gb_Click()
    LabelAnswer = "B"
    If LabelAnswer = Answer Then
        Call Form1.Win
    End If
    Unload Form3
End Sub
Private Sub gC_Click()
    LabelAnswer = "C"
    If LabelAnswer = Answer Then
        Call Form1.Win
    End If
    Unload Form3
End Sub
Private Sub gD_Click()
    LabelAnswer = "D"
    If LabelAnswer = Answer Then
        Call Form1.Win
    End If
    Unload Form3
End Sub

