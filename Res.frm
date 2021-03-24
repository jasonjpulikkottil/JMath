VERSION 5.00
Begin VB.Form Res 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   5535
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4455
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Tim 
      Interval        =   10
      Left            =   360
      Top             =   600
   End
   Begin VB.Label Mover 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   4695
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   6375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FFFF&
      Height          =   5535
      Left            =   0
      Top             =   0
      Width           =   4455
   End
   Begin VB.Label Blink 
      Caption         =   "Label3"
      Height          =   255
      Left            =   360
      TabIndex        =   12
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label8 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Speed"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   375
      Left            =   600
      TabIndex        =   11
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label SPEED 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   975
      Left            =   2400
      TabIndex        =   10
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label SECOND 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   375
      Left            =   2400
      TabIndex        =   9
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Time Taken"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Index           =   1
      Left            =   600
      TabIndex        =   8
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label PRCT 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2400
      TabIndex        =   7
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Accuracy"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   375
      Left            =   600
      TabIndex        =   6
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label CRT 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label TOT 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Correct Answers"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Index           =   0
      Left            =   600
      TabIndex        =   3
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Questions"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   1440
      MouseIcon       =   "Res.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF0000&
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   1440
      Shape           =   4  'Rounded Rectangle
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "Result"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "Res"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Form_Load()
Label1.ForeColor = RGB(255, 80, 85)
End Sub

Private Sub Label2_Click()
End
End Sub


Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim L1, L2, H1, H2, PosX, PosY

L1 = Label2.Left
H1 = Label2.Top

L2 = Label2.Left - -Label2.Width
H2 = Label2.Top - -Label2.Height

PosX = L1 - -X
PosY = H1 - -Y

If (PosX > L1 + 70 And PosX < L2 - 70) And (PosY > H1 + 70 And PosY < H2 - 70) Then
Blink.Caption = "YES"
Else
Blink.Caption = "NO"
End If

End Sub



Private Sub Mover_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = 1 Then
        Me.Left = Me.Left + X
        Me.Top = Me.Top + Y
    End If
End Sub

Private Sub Tim_Timer()
If Blink.Caption = "YES" Then
Label2.ForeColor = RGB(55, 255, 55)
Shape2.FillColor = RGB(20, 20, 180)
Else
Label2.ForeColor = RGB(255, 255, 255)
Shape2.FillColor = RGB(0, 0, 110)
End If
End Sub


