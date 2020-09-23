VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   Caption         =   "COMPUTER POLL"
   ClientHeight    =   4560
   ClientLeft      =   7020
   ClientTop       =   3420
   ClientWidth     =   4260
   LinkTopic       =   "Form2"
   ScaleHeight     =   4560
   ScaleWidth      =   4260
   Begin VB.CommandButton Command1 
      Caption         =   "CLOSE"
      Height          =   495
      Left            =   1440
      TabIndex        =   12
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Label lblDtop 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "d"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3240
      TabIndex        =   11
      Top             =   0
      Width           =   735
   End
   Begin VB.Label lblCtop 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "b"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   2280
      TabIndex        =   10
      Top             =   0
      Width           =   735
   End
   Begin VB.Label lblBtop 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "c"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1320
      TabIndex        =   9
      Top             =   0
      Width           =   735
   End
   Begin VB.Label lblAtop 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "a"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   0
      Width           =   735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   31
      X1              =   4080
      X2              =   240
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Label lbldgr 
      BackColor       =   &H00FF0000&
      Height          =   2415
      Left            =   3240
      TabIndex        =   2
      Top             =   600
      Width           =   615
   End
   Begin VB.Label lblcgr 
      BackColor       =   &H0000FF00&
      Height          =   2415
      Left            =   2280
      TabIndex        =   3
      Top             =   600
      Width           =   615
   End
   Begin VB.Label lblbgr 
      BackColor       =   &H000000FF&
      Height          =   2415
      Left            =   1320
      TabIndex        =   1
      Top             =   600
      Width           =   615
   End
   Begin VB.Label lblagr 
      BackColor       =   &H0000FFFF&
      Height          =   2415
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   615
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   30
      X1              =   4080
      X2              =   240
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   29
      X1              =   4080
      X2              =   240
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   28
      X1              =   4080
      X2              =   240
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   27
      X1              =   4080
      X2              =   240
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   26
      X1              =   4080
      X2              =   240
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   25
      X1              =   4080
      X2              =   240
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   24
      X1              =   4080
      X2              =   240
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   23
      X1              =   4080
      X2              =   240
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   22
      X1              =   4080
      X2              =   240
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   21
      X1              =   4080
      X2              =   240
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   20
      X1              =   4080
      X2              =   240
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   19
      X1              =   4080
      X2              =   240
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   18
      X1              =   4080
      X2              =   240
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   16
      X1              =   4080
      X2              =   4080
      Y1              =   360
      Y2              =   3840
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   15
      X1              =   480
      X2              =   480
      Y1              =   360
      Y2              =   3240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   14
      X1              =   720
      X2              =   720
      Y1              =   360
      Y2              =   3240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   13
      X1              =   960
      X2              =   960
      Y1              =   360
      Y2              =   3240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   12
      X1              =   1200
      X2              =   1200
      Y1              =   360
      Y2              =   3840
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   11
      X1              =   1440
      X2              =   1440
      Y1              =   360
      Y2              =   3240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   10
      X1              =   1680
      X2              =   1680
      Y1              =   360
      Y2              =   3240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   9
      X1              =   1920
      X2              =   1920
      Y1              =   360
      Y2              =   3240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   8
      X1              =   2160
      X2              =   2160
      Y1              =   360
      Y2              =   3840
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   7
      X1              =   2400
      X2              =   2400
      Y1              =   360
      Y2              =   3240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   6
      X1              =   2640
      X2              =   2640
      Y1              =   360
      Y2              =   3240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   5
      X1              =   2880
      X2              =   2880
      Y1              =   360
      Y2              =   3240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   4
      X1              =   3120
      X2              =   3120
      Y1              =   360
      Y2              =   3840
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   3
      X1              =   3360
      X2              =   3360
      Y1              =   360
      Y2              =   3240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   2
      X1              =   3600
      X2              =   3600
      Y1              =   360
      Y2              =   3240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   1
      X1              =   3840
      X2              =   3840
      Y1              =   360
      Y2              =   3240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   0
      X1              =   240
      X2              =   240
      Y1              =   360
      Y2              =   3840
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3240
      TabIndex        =   7
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   2280
      TabIndex        =   6
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "B"
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
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   3360
      Width           =   615
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Form2.Hide
    Unload Form2
End Sub

Private Sub Form_Load()
Dim Rand2 As Integer
Dim Rand3 As Integer
Dim Rand4 As Integer
Dim A As Integer
Dim B As Integer
Dim C As Integer
Dim D As Integer
    Randomize Timer
    If Form1.lblA.Visible = False Then
    Do While (D + B + C) <> 100
        D = Int(Rnd * 100) + 1
        Rand2 = 100 - D
        B = Int(Rnd * Rand2) + 1
        Rand3 = Rand2 - B
        C = Int(Rnd * 100) + 1
    Loop
    ElseIf Form1.lblB.Visible = False Then
    Do While (A + D + C) <> 100
        A = Int(Rnd * 100) + 1
        Rand2 = 100 - A
        D = Int(Rnd * Rand2) + 1
        Rand3 = Rand2 - D
        C = Int(Rnd * 100) + 1
    Loop
    ElseIf Form1.lblC.Visible = False Then
    Do While (A + D + C) <> 100
        A = Int(Rnd * 100) + 1
        Rand2 = 100 - A
        B = Int(Rnd * Rand2) + 1
        Rand3 = Rand2 - B
        C = Int(Rnd * 100) + 1
    Loop
    ElseIf Form1.lblD.Visible = False Then
    Do While (A + B + C) <> 100
        A = Int(Rnd * 100) + 1
        Rand2 = 100 - A
        B = Int(Rnd * Rand2) + 1
        Rand3 = Rand2 - B
        C = Int(Rnd * 100) + 1
    Loop
    ElseIf Form1.lblA.Visible = False And Form1.lblB.Visible = False Then
        Do While (C + D) <> 100
            If Answer = "C" Then
                C = Int(Rnd * 100) + 1
                Rand2 = 100 - C
                D = Int(Rnd * Rand2) + 1
            ElseIf Answer = "D" Then
                D = Int(Rnd * 100) + 1
                Rand2 = 100 - D
                C = Int(Rnd * Rand2) + 1
            End If
            A = 0
            B = 0
        Loop
    ElseIf Form1.lblC.Visible = False And Form1.lblB.Visible = False Then
        Do While (A + D) <> 100
            If Answer = "A" Then
                A = Int(Rnd * 100) + 1
                Rand2 = 100 - A
                D = Int(Rnd * Rand2) + 1
            ElseIf Answer = "D" Then
                D = Int(Rnd * 100) + 1
                Rand2 = 100 - D
                A = Int(Rnd * Rand2) + 1
            End If
            C = 0
            B = 0
        Loop
    ElseIf Form1.lblC.Visible = False And Form1.lblA.Visible = False Then
        Do While (B + D) <> 100
            If Answer = "B" Then
                B = Int(Rnd * 100) + 1
                Rand2 = 100 - B
                D = Int(Rnd * Rand2) + 1
            ElseIf Answer = "D" Then
                D = Int(Rnd * 100) + 1
                Rand2 = 100 - D
                B = Int(Rnd * Rand2) + 1
            End If
            C = 0
            A = 0
        Loop
    ElseIf Form1.lblA.Visible = False And Form1.lblD.Visible = False Then
        Do While (B + C) <> 100
            If Answer = "B" Then
                B = Int(Rnd * 100) + 1
                Rand2 = 100 - B
                C = Int(Rnd * Rand2) + 1
            ElseIf Answer = "C" Then
                C = Int(Rnd * 100) + 1
                Rand2 = 100 - C
                B = Int(Rnd * Rand2) + 1
            End If
            A = 0
            D = 0
        Loop
    ElseIf Form1.lblD.Visible = False And Form1.lblB.Visible = False Then
        Do While (A + C) <> 100
            If Answer = "A" Then
                A = Int(Rnd * 100) + 1
                Rand2 = 100 - A
                C = Int(Rnd * Rand2) + 1
            ElseIf Answer = "C" Then
                C = Int(Rnd * 100) + 1
                Rand2 = 100 - C
                A = Int(Rnd * Rand2) + 1
            End If
            D = 0
            B = 0
        Loop
    ElseIf Form1.lblC.Visible = False And Form1.lblD.Visible = False Then
        Do While (A + B) <> 100
            If Answer = "A" Then
                A = Int(Rnd * 100) + 1
                Rand2 = 100 - A
                B = Int(Rnd * Rand2) + 1
            ElseIf Answer = "B" Then
                B = Int(Rnd * 100) + 1
                Rand2 = 100 - B
                A = Int(Rnd * Rand2) + 1
            End If
            C = 0
            D = 0
        Loop
    Else
    If Answer = "A" Then
    Do While (A + B + C + D) <> 100
        A = Int(Rnd * 100) + 1
        Rand2 = 100 - A
        B = Int(Rnd * Rand2) + 1
        Rand3 = Rand2 - B
        C = Int(Rnd * 100) + 1
        Rand4 = 100 - C
        D = Int(Rnd * Rand4) + 1
    Loop
    ElseIf Answer = "B" Then
    Do While (A + B + C + D) <> 100
        B = Int(Rnd * 100) + 1
        Rand2 = 100 - A
        A = Int(Rnd * Rand2) + 1
        Rand3 = Rand2 - B
        C = Int(Rnd * 100) + 1
        Rand4 = 100 - C
        D = Int(Rnd * Rand4) + 1
    Loop
    ElseIf Answer = "C" Then
    Do While (A + B + C + D) <> 100
        C = Int(Rnd * 100) + 1
        Rand2 = 100 - A
        B = Int(Rnd * Rand2) + 1
        Rand3 = Rand2 - B
        D = Int(Rnd * 100) + 1
        Rand4 = 100 - C
        A = Int(Rnd * Rand4) + 1
    Loop
    ElseIf Answer = "D" Then
    Do While (A + B + C + D) <> 100
        D = Int(Rnd * 100) + 1
        Rand2 = 100 - A
        B = Int(Rnd * Rand2) + 1
        Rand3 = Rand2 - B
        C = Int(Rnd * 100) + 1
        Rand4 = 100 - C
        A = Int(Rnd * Rand4) + 1
    Loop
    End If
    End If
    Let lblagr.Height = lblagr.Height * (A / 100)
    Let lblbgr.Height = lblbgr.Height * (B / 100)
    Let lblcgr.Height = lblcgr.Height * (C / 100)
    Let lbldgr.Height = lbldgr.Height * (D / 100)
    
    Let lblAtop.Caption = "% " & A
    Let lblBtop.Caption = "% " & B
    Let lblCtop.Caption = "% " & C
    Let lblDtop.Caption = "% " & D
End Sub
