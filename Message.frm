VERSION 5.00
Begin VB.Form Message 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   " If Button = 1 Then"
   ClientHeight    =   2670
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4455
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Tim2 
      Interval        =   10
      Left            =   3960
      Top             =   1560
   End
   Begin VB.Timer Tim1 
      Interval        =   10
      Left            =   120
      Top             =   1560
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2
      Left            =   3840
      Top             =   240
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
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.Label Blink2 
      Caption         =   "Label4"
      Height          =   255
      Left            =   2640
      TabIndex        =   4
      Top             =   1200
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Blink1 
      Caption         =   "Label4"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   1200
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H000000FF&
      Height          =   2655
      Left            =   0
      Top             =   0
      Width           =   4455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "No"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2400
      MouseIcon       =   "Message.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   2400
      Shape           =   4  'Rounded Rectangle
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Yes"
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
      MouseIcon       =   "Message.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF0000&
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   600
      Shape           =   4  'Rounded Rectangle
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
      Caption         =   "Are you sure to quit ?"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   600
      Width           =   2775
   End
End
Attribute VB_Name = "message"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Label2_Click()


Me.Hide
Timer1.Enabled = True



End Sub



Private Sub Label3_Click()
Me.Hide

End Sub



Private Sub Mover_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = 1 Then
        Me.Left = Me.Left + X
        Me.Top = Me.Top + Y
    End If
End Sub

Private Sub Timer1_Timer()
End
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

If KeyAscii = 121 Or KeyAscii = 89 Or KeyAscii = 13 Then End

If KeyAscii = 78 Or KeyAscii = 110 Then Me.Hide

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
Blink1.Caption = "YES"
Else
Blink1.Caption = "NO"
End If

End Sub


Private Sub Tim1_Timer()
If Blink1.Caption = "YES" Then
Label2.ForeColor = RGB(255, 250, 50)
Shape2.FillColor = RGB(20, 20, 180)
Else
Label2.ForeColor = RGB(250, 50, 50)
Shape2.FillColor = RGB(0, 0, 110)
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

If (PosX > L1 + 70 And PosX < L2 - 70) And (PosY > H1 + 70 And PosY < H2 - 70) Then
Blink2.Caption = "YES"
Else
Blink2.Caption = "NO"
End If

End Sub


Private Sub Tim2_Timer()
If Blink2.Caption = "YES" Then
Label3.ForeColor = RGB(255, 250, 50)
Shape1.FillColor = RGB(20, 20, 180)
Else
Label3.ForeColor = RGB(250, 50, 50)
Shape1.FillColor = RGB(0, 0, 110)
End If
End Sub











