VERSION 5.00
Begin VB.Form Starter 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "Starter"
   ClientHeight    =   4830
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5070
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Tim2 
      Interval        =   10
      Left            =   4200
      Top             =   3120
   End
   Begin VB.Timer Tim1 
      Interval        =   10
      Left            =   4200
      Top             =   1800
   End
   Begin VB.Timer Tim 
      Interval        =   10
      Left            =   4080
      Top             =   120
   End
   Begin VB.Label Label5 
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
      Height          =   1455
      Left            =   -120
      TabIndex        =   8
      Top             =   3960
      Visible         =   0   'False
      Width           =   6375
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
      Height          =   1455
      Left            =   -1920
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   6375
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00808080&
      Height          =   390
      Left            =   4560
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   4560
      TabIndex        =   6
      ToolTipText     =   "Close"
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Blink2 
      Caption         =   "Label5"
      Height          =   135
      Left            =   480
      TabIndex        =   5
      Top             =   3000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Blink1 
      Caption         =   "Label5"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   1800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Blink 
      Caption         =   "Label5"
      Height          =   255
      Left            =   3240
      TabIndex        =   3
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Take the test"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   735
      Left            =   1200
      MouseIcon       =   "Starter.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   3120
      Width           =   2535
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FF00&
      FillColor       =   &H00004000&
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   1080
      Shape           =   4  'Rounded Rectangle
      Top             =   2880
      Width           =   2895
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Practise Mode"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   615
      Left            =   1080
      MouseIcon       =   "Starter.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   1920
      Width           =   2895
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0080FF80&
      FillColor       =   &H00004000&
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   1080
      Shape           =   4  'Rounded Rectangle
      Top             =   1680
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
      Caption         =   "Please select an option"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   600
      Width           =   3015
   End
End
Attribute VB_Name = "Starter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Form1.TCalc
End Sub

Private Sub Label2_Click()
Me.Tag = 1




Form1.Qno.Visible = False
Form1.Shape9.Visible = False
Form1.Shape10.Visible = False
Form1.Hider.Visible = False
Form1.Label14.Visible = False
Form1.Shape11.Visible = False





Form1.Show
Me.Hide
End Sub

Private Sub Label3_Click()
Dim J As Integer
Me.Tag = 2


Form1.Label12.Visible = True
Form1.Label15.Visible = True

Form1.Shape12.Visible = True


Form1.Shape7.Visible = False
Form1.Label5.Visible = False
Form1.Label11.Visible = False


J = 2000
Form1.Check.Caption = "Submit"
Form1.Shape6.Left = Form1.Shape6.Left + J
Form1.Check.Left = Form1.Check.Left + J
Form1.Label10.Left = Form1.Label10.Left + J



Form1.Show
Me.Hide

End Sub

Private Sub Label4_Click()

message.Show
End Sub





Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim L1, L2, H1, H2, PosX, PosY

L1 = Label4.Left
H1 = Label4.Top

L2 = Label4.Left - -Label4.Width
H2 = Label4.Top - -Label4.Height

PosX = L1 - -X
PosY = H1 - -Y

If (PosX > L1 + 50 And PosX < L2 - 50) And (PosY > H1 + 50 And PosY < H2 - 50) Then
Blink.Caption = "YES"
Else
Blink.Caption = "NO"
End If

End Sub






Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
        Me.Left = Me.Left + X
        Me.Top = Me.Top + Y
    End If
End Sub

Private Sub Mover_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox Button
If Button = 1 Then
        Me.Left = Me.Left + X
        Me.Top = Me.Top + Y
    End If
End Sub

Private Sub Tim_Timer()
If Blink.Caption = "YES" Then
Label4.ForeColor = RGB(255, 250, 50)
Shape4.BorderColor = RGB(150, 150, 150)

Else
Label4.ForeColor = RGB(250, 50, 50)

Shape4.BorderColor = RGB(100, 90, 90)
End If
End Sub






Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim L1, L2, H1, H2, PosX, PosY

L1 = Label2.Left
H1 = Label2.Top

L2 = Label2.Left - -Label2.Width
H2 = Label2.Top - -Label2.Height

PosX = L1 - -X
PosY = H1 - -Y

If (PosX > L1 + 90 And PosX < L2 - 110) And (PosY > H1 + 110 And PosY < H2 - 110) Then
Blink1.Caption = "YES"
Else
Blink1.Caption = "NO"
End If

End Sub


Private Sub Tim1_Timer()
If Blink1.Caption = "YES" Then

Shape2.FillColor = RGB(20, 120, 30)
Else

Shape2.FillColor = RGB(0, 90, 0)
End If
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim L1, L2, H1, H2, PosX, PosY

L1 = Label3.Left
H1 = Label3.Top

L2 = Label3.Left - -Label3.Width
H2 = Label3.Top - -Label3.Height

PosX = L1 - -X
PosY = H1 - -Y

If (PosX > L1 + 90 And PosX < L2 - 110) And (PosY > H1 + 110 And PosY < H2 - 110) Then
Blink2.Caption = "YES"
Else
Blink2.Caption = "NO"
End If

End Sub


Private Sub Tim2_Timer()
If Blink2.Caption = "YES" Then

Shape1.FillColor = RGB(20, 120, 30)
Else

Shape1.FillColor = RGB(0, 90, 0)
End If
End Sub

