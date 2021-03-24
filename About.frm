VERSION 5.00
Begin VB.Form About 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   2640
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3960
   LinkTopic       =   "Form2"
   ScaleHeight     =   2640
   ScaleWidth      =   3960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Tim 
      Interval        =   10
      Left            =   3480
      Top             =   480
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
      Height          =   1815
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Version 3.1"
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
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   960
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "TIP : You can increase your calculation skills by regular use of this tool."
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   615
      Left            =   360
      TabIndex        =   3
      Top             =   1200
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Label Blink 
      Caption         =   "Label3"
      Height          =   255
      Left            =   3480
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   960
      MouseIcon       =   "About.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Shape shape1 
      BorderColor     =   &H000000FF&
      Height          =   2640
      Left            =   0
      Top             =   0
      Width           =   3960
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF0000&
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   1200
      Shape           =   4  'Rounded Rectangle
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Created by Jason J Pulikkottil"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   3255
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 32 Then Me.Hide
End Sub

Private Sub Label2_Click()
Me.Hide
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
Label2.ForeColor = RGB(255, 250, 50)
Shape2.FillColor = RGB(20, 20, 180)
Else
Label2.ForeColor = RGB(250, 50, 50)
Shape2.FillColor = RGB(0, 0, 110)
End If
End Sub

