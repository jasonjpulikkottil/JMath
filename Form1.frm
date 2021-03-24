VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4410
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6855
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Segoe UI Semibold"
      Size            =   9
      Charset         =   0
      Weight          =   600
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MouseIcon       =   "Form1.frx":2FBA
   ScaleHeight     =   4410
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "1"
   Begin VB.Timer Timer3 
      Interval        =   600
      Left            =   5880
      Top             =   3840
   End
   Begin VB.TextBox TT1 
      Height          =   330
      Left            =   3600
      TabIndex        =   29
      Text            =   "0"
      Top             =   2280
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Timer Tim1 
      Interval        =   10
      Left            =   1080
      Top             =   1800
   End
   Begin VB.Timer Tim 
      Interval        =   10
      Left            =   5880
      Top             =   120
   End
   Begin VB.TextBox Rslt 
      Height          =   330
      Left            =   4440
      TabIndex        =   25
      Text            =   "0"
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Timer Progress 
      Enabled         =   0   'False
      Interval        =   70
      Left            =   1680
      Tag             =   "0"
      Top             =   120
   End
   Begin VB.TextBox Text 
      Height          =   288
      Left            =   1053
      TabIndex        =   16
      Text            =   "X"
      Top             =   4095
      Visible         =   0   'False
      Width           =   126
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6480
      Top             =   3360
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6480
      Top             =   2880
   End
   Begin VB.TextBox Ans 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   24
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   4440
      MaxLength       =   7
      TabIndex        =   8
      Top             =   1800
      Width           =   1815
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
      Height          =   615
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Visible         =   0   'False
      Width           =   6375
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(Tab)"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   345
      Left            =   1560
      MouseIcon       =   "Form1.frx":32C4
      MousePointer    =   99  'Custom
      TabIndex        =   31
      Top             =   3720
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   6240
      MouseIcon       =   "Form1.frx":35CE
      MousePointer    =   99  'Custom
      TabIndex        =   30
      Top             =   3840
      Width           =   615
   End
   Begin VB.Label Blink1 
      Caption         =   "Label13"
      Height          =   255
      Left            =   1080
      TabIndex        =   28
      Top             =   2400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Blink 
      Caption         =   "Label13"
      Height          =   255
      Left            =   4920
      TabIndex        =   27
      Top             =   2760
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Show Result"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   1800
      MouseIcon       =   "Form1.frx":38D8
      MousePointer    =   99  'Custom
      TabIndex        =   26
      Top             =   3480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Shape Shape12 
      BorderColor     =   &H00FF0000&
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   1560
      Shape           =   4  'Rounded Rectangle
      Top             =   3240
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   705
      Left            =   1680
      MouseIcon       =   "Form1.frx":3BE2
      MousePointer    =   99  'Custom
      TabIndex        =   24
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Shape Shape11 
      BorderColor     =   &H0000FF00&
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   1560
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Shape Hider 
      FillStyle       =   0  'Solid
      Height          =   1455
      Left            =   1200
      Top             =   1680
      Width           =   5535
   End
   Begin VB.Shape Shape10 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   1080
      Top             =   1440
      Width           =   15
   End
   Begin VB.Label Qno 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   240
      MouseIcon       =   "Form1.frx":3EEC
      MousePointer    =   99  'Custom
      TabIndex        =   23
      Top             =   360
      Width           =   495
   End
   Begin VB.Shape Shape9 
      BorderColor     =   &H00C00000&
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   240
      Shape           =   3  'Circle
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(Space)"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   345
      Left            =   3840
      MouseIcon       =   "Form1.frx":41F6
      MousePointer    =   99  'Custom
      TabIndex        =   22
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(Enter)"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   345
      Left            =   1680
      MouseIcon       =   "Form1.frx":4500
      MousePointer    =   99  'Custom
      TabIndex        =   21
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label Opl4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   240
      MouseIcon       =   "Form1.frx":480A
      MousePointer    =   99  'Custom
      TabIndex        =   20
      Top             =   3120
      Width           =   495
   End
   Begin VB.Shape Op4 
      BorderColor     =   &H00FF0000&
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   240
      Shape           =   5  'Rounded Square
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "!"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   360
      Left            =   6195
      MouseIcon       =   "Form1.frx":4B14
      MousePointer    =   99  'Custom
      TabIndex        =   19
      Top             =   3890
      Width           =   600
   End
   Begin VB.Shape Shape8 
      BorderColor     =   &H00FF0000&
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   6315
      Shape           =   5  'Rounded Square
      Top             =   3855
      Width           =   375
   End
   Begin VB.Label Opl3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   240
      MouseIcon       =   "Form1.frx":4E1E
      MousePointer    =   99  'Custom
      TabIndex        =   18
      Top             =   1725
      Width           =   495
   End
   Begin VB.Shape Op3 
      BorderColor     =   &H00FF0000&
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   240
      Shape           =   5  'Rounded Square
      Top             =   1725
      Width           =   495
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      X1              =   930
      X2              =   930
      Y1              =   0
      Y2              =   6960
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "J Math"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   477
      Left            =   2808
      TabIndex        =   17
      Top             =   117
      Width           =   1530
   End
   Begin VB.Label Opl2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   240
      MouseIcon       =   "Form1.frx":5128
      MousePointer    =   99  'Custom
      TabIndex        =   15
      Top             =   1065
      Width           =   495
   End
   Begin VB.Shape Op2 
      BorderColor     =   &H00FF0000&
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   240
      Shape           =   5  'Rounded Square
      Top             =   1065
      Width           =   495
   End
   Begin VB.Label Opl1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   240
      MouseIcon       =   "Form1.frx":5432
      MousePointer    =   99  'Custom
      TabIndex        =   14
      Top             =   2475
      Width           =   495
   End
   Begin VB.Shape op1 
      BorderColor     =   &H00FF0000&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   240
      Shape           =   5  'Rounded Square
      Top             =   2355
      Width           =   495
   End
   Begin VB.Label Result2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Correct Answer Is "
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   600
      TabIndex        =   13
      Top             =   5400
      Visible         =   0   'False
      Width           =   5655
   End
   Begin VB.Label Result1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Correct Answer"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   615
      Left            =   1800
      TabIndex        =   12
      Top             =   4560
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Next"
      ForeColor       =   &H0080FF80&
      Height          =   495
      Left            =   3840
      MouseIcon       =   "Form1.frx":573C
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H00FF0000&
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   3840
      Shape           =   4  'Rounded Rectangle
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label Check 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Check Answer"
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   1680
      MouseIcon       =   "Form1.frx":5A46
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H00FF0000&
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   1560
      Shape           =   4  'Rounded Rectangle
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label Label9 
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
      Left            =   6363
      TabIndex        =   9
      ToolTipText     =   "Close"
      Top             =   117
      Width           =   378
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00808080&
      Height          =   395
      Left            =   6360
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   24
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   738
      Left            =   3744
      TabIndex        =   7
      Top             =   1755
      Width           =   495
   End
   Begin VB.Label Op 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   24
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   738
      Left            =   2223
      TabIndex        =   6
      Top             =   1755
      Width           =   495
   End
   Begin VB.Label Num2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "234"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   24
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   738
      Left            =   2691
      TabIndex        =   5
      Top             =   1755
      Width           =   1098
   End
   Begin VB.Label Num1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "234"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   24
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   738
      Left            =   1170
      TabIndex        =   4
      Top             =   1755
      Width           =   1098
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Level 4"
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   5400
      MouseIcon       =   "Form1.frx":5D50
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FF0000&
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   5400
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Level 3"
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   3960
      MouseIcon       =   "Form1.frx":605A
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FF0000&
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   3960
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Level 2"
      ForeColor       =   &H0000FFFF&
      Height          =   378
      Left            =   2520
      MouseIcon       =   "Form1.frx":6364
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   837
      Width           =   1215
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF0000&
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   2520
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Level 1"
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   1080
      MouseIcon       =   "Form1.frx":666E
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   840
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   1080
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Function TCalc()

Dim Sec, Min, Hour

Hour = Val(Mid(Time, 1, 2))
Min = Val(Mid(Time, 4, 2))
Sec = Val(Mid(Time, 7, 2))

TCalc = Val(Hour * 60 * 60 + Min * 60 + Sec)



End Function


Public Function Num(Digits As Integer)
Pwr = 10 ^ Digits
Randomize

i = Int(Rnd * Pwr)
Num = i
End Function

Public Function Calc()
If UCase(Op.Caption) = "X" Then Calc = Val(Num1) * Val(Num2)
If UCase(Op.Caption) = "+" Then Calc = Val(Num1) + Val(Num2)
If UCase(Op.Caption) = "-" Then Calc = Val(Num1) - Val(Num2)
If UCase(Op.Caption) = "/" Then Calc = Val(Num1) / Val(Num2)



End Function

Public Function Cn()

i = Val(Me.Tag)
If i = 1 Then
N1 = 1
N2 = 1
End If
If i = 2 Then
N1 = 2
N2 = 1
End If
If i = 3 Then
N1 = 2
N2 = 2
End If
If i = 4 Then
N1 = 3
N2 = 2
End If
Num1.Caption = Num(Val(N1))
Num2.Caption = Num(Val(N2))

ONe:

If Text.Text = "/" Then
Num2.Caption = Num(Val(N2))
Num1.Caption = Val(Num2.Caption) * Num(2)
End If

If Text.Text = "/" And Num2 = 0 And Num1 = 0 Then GoTo ONe

Op.Caption = Text.Text

Ans.Text = ""
Timer2.Enabled = True
End Function

Public Function RESULTShow()
If (Val(Qno.Caption) - 1) <> 0 Then
Res.TOT = Val(Qno.Caption) - 1
Res.CRT = Rslt.Text
Res.PRCT = Round((Val(Res.CRT) / Val(Res.TOT)) * 100, 3) & " %"

tmnw = TCalc

Res.SECOND.Caption = Abs(Val(tmnw) - Val(TT1.Text)) & " Seconds"
Res.SPEED.Caption = Round((Val(Res.TOT) / Val(Res.SECOND.Caption)) * 60, 2) & " Calculations Per Minute"

Else
 Res.TOT.Caption = "Not started"
 Res.CRT.Caption = "Not started"
Res.PRCT.Caption = "Not started"
 Res.SECOND.Caption = "Not started"
Res.SPEED.Caption = "Not started"






End If

Me.Hide
Res.Show

End Function

Public Function Submit()


If Val(Starter.Tag) = 2 Then

If Calc = Val(Ans.Text) And Ans.Text <> "" Then Rslt = Val(Rslt) + 1

Cn

Qno = Val(Qno.Caption) + 1
Shape10.Width = 15
Progress.Tag = 0




Else

If Ans.Text = "" Then
About.Label1.Caption = "Please Input the Answer"

About.Label3.Visible = False
About.Label4.Visible = False



About.Show
Else

If Calc = Val(Ans.Text) Then

Result1.Visible = True
Result1.Caption = "Correct Answer"
Result1.ForeColor = RGB(50, 255, 50)
Result2.Visible = False

Else

Result1.Visible = True
Result1.Caption = "Wrong Answer"
Result1.ForeColor = RGB(255, 50, 50)
Result2.Visible = True

Result2.Caption = "Correct Answer Is " & Calc

End If

Timer1.Enabled = True

End If
End If

End Function








Private Sub Ans_Change()
If Ans.Text <> "-" And Ans.Text <> "0" And Val(Ans.Text) = 0 Then
Ans.Text = ""
End If
End Sub






Private Sub Ans_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Submit
End If


If KeyAscii = 32 Then

If Val(Starter.Tag) = 1 Then
Timer2.Enabled = True
Cn
End If
End If


If KeyAscii = 9 Then

If Val(Starter.Tag) = 2 Then

RESULTShow

End If
End If




End Sub

Private Sub Check_Click()
Submit
End Sub



Private Sub Form_Load()
Cn

End Sub


Private Sub Label1_Click()
Cn
Shape1.FillColor = &HFF0000
Shape2.FillColor = &H800000
Shape3.FillColor = &H800000
Shape4.FillColor = &H800000
Me.Tag = 1
End Sub

Private Sub Label10_Click()
Submit
End Sub

Private Sub Label11_Click()
Timer2.Enabled = True
Cn

End Sub

Private Sub Label12_Click()
RESULTShow
End Sub

Private Sub Label13_Click()
About.Label1.Caption = "Created by Jason J Pulikkottil"
About.Label3.Visible = True
About.Label4.Visible = True

About.Show

End Sub

Private Sub Label14_Click()

TT1.Text = TCalc


Progress.Enabled = True
Hider.Visible = False
Label14.Visible = False
Shape11.Visible = False

End Sub



Private Sub Label15_Click()
RESULTShow
End Sub

Private Sub Label2_Click()
Cn
Shape1.FillColor = &H800000
Shape2.FillColor = &HFF0000
Shape3.FillColor = &H800000
Shape4.FillColor = &H800000
Me.Tag = 2
End Sub

Private Sub Label3_Click()
Cn
Shape1.FillColor = &H800000
Shape2.FillColor = &H800000
Shape3.FillColor = &HFF0000
Shape4.FillColor = &H800000
Me.Tag = 3
End Sub

Private Sub Label4_Click()
Cn
Shape1.FillColor = &H800000
Shape2.FillColor = &H800000
Shape3.FillColor = &H800000
Shape4.FillColor = &HFF0000
Me.Tag = 4
End Sub

Private Sub Label5_Click()
Timer2.Enabled = True
Cn

End Sub

Private Sub Label7_Click()
About.Label1.Caption = "Created by Jason J Pulikkottil"
About.Show
End Sub



Private Sub Label9_Click()
message.Show
End Sub




Private Sub Mover_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = 1 Then
        Me.Left = Me.Left + X
        Me.Top = Me.Top + Y
    End If
End Sub

Private Sub Opl1_Click()

op1.FillColor = &HFF0000
Op2.FillColor = &H800000
Op3.FillColor = &H800000
Op4.FillColor = &H800000
Text.Text = "X"
Cn
End Sub

Private Sub Opl2_Click()
op1.FillColor = &H800000
Op2.FillColor = &HFF0000
Op3.FillColor = &H800000
Op4.FillColor = &H800000
Text.Text = "+"
Cn
End Sub

Private Sub Opl3_Click()
op1.FillColor = &H800000
Op2.FillColor = &H800000
Op3.FillColor = &HFF0000
Op4.FillColor = &H800000
Text.Text = "-"
Cn
End Sub

Private Sub Opl4_Click()

op1.FillColor = &H800000
Op2.FillColor = &H800000
Op3.FillColor = &H800000
Op4.FillColor = &HFF0000
Text.Text = "/"
Cn
End Sub



Private Sub Timer1_Timer()

Timer2.Enabled = False

If Form1.Height <= 6161 Then
Form1.Height = Form1.Height + 60

Else
Timer1.Enabled = False
End If


End Sub

Private Sub Timer2_Timer()
If Me.Height >= 4488 Then
Me.Height = Me.Height - 60

Else
Timer2.Enabled = False
End If
End Sub

Private Sub Progress_Timer()
On Error Resume Next

If Shape1.FillColor = &HFF0000 Then
Progress.Interval = 20
Else
Progress.Interval = 75
End If




If Shape10.Width < 5535 Then

Shape10.Width = Shape10.Width + 20

Shape10.FillColor = RGB(30 + Int(Val(Progress.Tag)), 255 - Int(Val(Progress.Tag)), 30)
Progress.Tag = Progress.Tag + 0.8


Else



Qno = Val(Qno.Caption) + 1
Shape10.Width = 15
Progress.Tag = 0

Timer2.Enabled = True
Cn

End If
End Sub






Private Sub Label9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim L1, L2, H1, H2, PosX, PosY

L1 = Label9.Left
H1 = Label9.Top

L2 = Label9.Left - -Label9.Width
H2 = Label9.Top - -Label9.Height

PosX = L1 - -X
PosY = H1 - -Y

If (PosX > L1 + 50 And PosX < L2 - 50) And (PosY > H1 + 50 And PosY < H2 - 50) Then
Blink.Caption = "YES"
Else
Blink.Caption = "NO"
End If

End Sub


Private Sub Tim_Timer()
If Blink.Caption = "YES" Then
Label9.ForeColor = RGB(255, 50, 250)
Shape5.BorderColor = RGB(150, 150, 150)

Else
Label9.ForeColor = RGB(250, 50, 50)

Shape5.BorderColor = RGB(100, 90, 90)
End If
End Sub








Private Sub Label14_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim L1, L2, H1, H2, PosX, PosY

L1 = Label14.Left
H1 = Label14.Top

L2 = Label14.Left - -Label14.Width
H2 = Label14.Top - -Label14.Height

PosX = L1 - -X
PosY = H1 - -Y

If (PosX > L1 + 90 And PosX < L2 - 90) And (PosY > H1 + 90 And PosY < H2 - 90) Then
Blink1.Caption = "YES"
Else
Blink1.Caption = "NO"
End If

End Sub


Private Sub Tim1_Timer()
If Blink1.Caption = "YES" Then

Shape11.FillColor = RGB(20, 120, 30)
Else

Shape11.FillColor = RGB(0, 90, 0)
End If
End Sub

Private Sub Timer3_Timer()
If Label7.Visible = True Then

Label7.Visible = False

Else

Label7.Visible = True

End If

End Sub
