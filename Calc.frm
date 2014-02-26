VERSION 5.00
Begin VB.Form Calc 
   BackColor       =   &H80000012&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calc"
   ClientHeight    =   4470
   ClientLeft      =   2865
   ClientTop       =   2175
   ClientWidth     =   4350
   FillColor       =   &H000000FF&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   4350
   Begin VB.CommandButton point 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   17
      Left            =   3480
      TabIndex        =   17
      Top             =   3600
      Width           =   735
   End
   Begin VB.CommandButton change 
      Caption         =   "+-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   16
      Left            =   3480
      TabIndex        =   16
      Top             =   2760
      Width           =   735
   End
   Begin VB.CommandButton delete2 
      Caption         =   "CE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   15
      Left            =   3480
      TabIndex        =   15
      Top             =   1920
      Width           =   735
   End
   Begin VB.CommandButton equal 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   14
      Left            =   3480
      TabIndex        =   14
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton sign4 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   13
      Left            =   2640
      TabIndex        =   13
      Top             =   3600
      Width           =   735
   End
   Begin VB.CommandButton sign3 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   12
      Left            =   2640
      TabIndex        =   12
      Top             =   2760
      Width           =   735
   End
   Begin VB.CommandButton sign2 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   11
      Left            =   2640
      TabIndex        =   11
      Top             =   1920
      Width           =   735
   End
   Begin VB.CommandButton sign1 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   10
      Left            =   2640
      TabIndex        =   10
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton number0 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   3600
      Width           =   2415
   End
   Begin VB.CommandButton number9 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   9
      Left            =   1800
      TabIndex        =   9
      Top             =   2760
      Width           =   735
   End
   Begin VB.CommandButton number6 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   6
      Left            =   1800
      TabIndex        =   6
      Top             =   1920
      Width           =   735
   End
   Begin VB.CommandButton number3 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   3
      Left            =   1800
      TabIndex        =   3
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton number8 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   8
      Left            =   960
      TabIndex        =   8
      Top             =   2760
      Width           =   735
   End
   Begin VB.CommandButton number7 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   7
      Left            =   120
      TabIndex        =   7
      Top             =   2760
      Width           =   735
   End
   Begin VB.CommandButton number5 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   5
      Left            =   960
      TabIndex        =   5
      Top             =   1920
      Width           =   735
   End
   Begin VB.CommandButton number4 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   735
   End
   Begin VB.CommandButton number2 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   2
      Left            =   960
      TabIndex        =   2
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton number1 
      BackColor       =   &H80000008&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   120
      MaskColor       =   &H0000FFFF&
      TabIndex        =   1
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox text1 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   855
      Left            =   120
      TabIndex        =   18
      Text            =   "0"
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "Calc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Double
Dim b As Double
Dim s As String
Dim buf1 As String
Dim buf2 As String
Dim buf As Double
Dim point1 As Boolean
Dim flag As Boolean
Dim number As Boolean
Dim zero As Boolean
Dim die As Boolean
Dim change1 As Boolean
Dim equal1 As Boolean
Dim full As Boolean

Private Sub change_Click(Index As Integer)
    If change1 Then
        buf = text1.Text
        buf = -buf
        text1.Text = buf
    End If
End Sub


Private Sub delete2_Click(Index As Integer)
    text1.Text = "0"
    point1 = False
    flag = False
    number = False
    die = False
    change1 = False
    full = False
End Sub

Private Sub equal_Click(Index As Integer)
    If (equal1) Then
        b = text1.Text
    End If
    If s = "+" Then
        text1.Text = a + b
    ElseIf s = "-" Then
        text1.Text = a - b
    ElseIf s = "*" Then
        text1.Text = a * b
    ElseIf s = "/" Then
        If b = 0 Then
            text1.Text = "You die, bustard!!"
            die = True
        Else
            text1.Text = a / b
        End If
    End If
    equal1 = False
    flag = True
    number = False
    point1 = False
    zero = False
    full = False
    If Not (die) Then
        a = text1.Text
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    text1.SetFocus
    Debug.Print KeyAscii
    If (Chr(KeyAscii) = "0") Then
        number0_Click (0)
    ElseIf (Chr(KeyAscii) = "1") Then
        number1_Click (0)
    ElseIf (Chr(KeyAscii) = "2") Then
        number2_Click (0)
    ElseIf (Chr(KeyAscii) = "3") Then
        number3_Click (0)
    ElseIf (Chr(KeyAscii) = "4") Then
        number4_Click (0)
    ElseIf (Chr(KeyAscii) = "5") Then
        number5_Click (0)
    ElseIf (Chr(KeyAscii) = "6") Then
        number6_Click (0)
    ElseIf (Chr(KeyAscii) = "7") Then
        number7_Click (0)
    ElseIf (Chr(KeyAscii) = "8") Then
        number8_Click (0)
    ElseIf (Chr(KeyAscii) = "9") Then
        number9_Click (0)
    ElseIf (Chr(KeyAscii) = "+") Then
        sign1_Click (0)
    ElseIf (Chr(KeyAscii) = "-") Then
        sign2_Click (0)
    ElseIf (Chr(KeyAscii) = "*") Then
        sign3_Click (0)
    ElseIf (Chr(KeyAscii) = "/") Then
        sign4_Click (0)
    ElseIf (Chr(KeyAscii) = "," Or KeyAscii = 110) Then
        point_Click (0)
    ElseIf (Chr(KeyAscii) = "=" Or KeyAscii = 13) Then
        equal_Click (0)
    ElseIf (KeyAscii = 8 Or KeyAscii = 46) Then
        delete2_Click (0)
    End If
    KeyAscii = 0
End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 120 Then
        change_Click (0)
    End If
    KeyCode = 0
End Sub

Private Sub number0_Click(Index As Integer)
    zero = True
    buf1 = text1.Text
    buf = CDbl(buf1)
    buf2 = buf
    point2 = Right(buf1, 1)
    If (buf1 <> buf2 And point2 <> ",") Then
        full = True
        text1.Text = Left(text1.Text, Len(text1.Text) - 1)
    End If
    If Not (full) Then
        If (number) Then
            text1.Text = text1.Text & "0"
        End If
        If (flag) Then
            text1.Text = "0"
        End If
    End If
End Sub

Private Sub number1_Click(Index As Integer)
buf1 = text1.Text
    buf = CDbl(buf1)
    buf2 = buf
    point2 = Right(buf1, 1)
    If (buf1 <> buf2 And point2 <> ",") Then
        full = True
        text1.Text = Left(text1.Text, Len(text1.Text) - 1)
    End If
    If Not (full) Then
        If (number) Then
            text1.Text = text1.Text & "1"
        Else
            text1.Text = "1"
            number = True
        End If
    End If
    change1 = True
End Sub

Private Sub number2_Click(Index As Integer)
    buf1 = text1.Text
    buf = CDbl(buf1)
    buf2 = buf
    point2 = Right(buf1, 1)
    If (buf1 <> buf2 And point2 <> ",") Then
        full = True
        text1.Text = Left(text1.Text, Len(text1.Text) - 1)
    End If
    If Not (full) Then
        If (number) Then
            text1.Text = text1.Text & "2"
        Else
            text1.Text = "2"
            number = True
        End If
    End If
    change1 = True
End Sub

Private Sub number3_Click(Index As Integer)
    buf1 = text1.Text
    buf = CDbl(buf1)
    buf2 = buf
    point2 = Right(buf1, 1)
    If (buf1 <> buf2 And point2 <> ",") Then
        full = True
        text1.Text = Left(text1.Text, Len(text1.Text) - 1)
    End If
    If Not (full) Then
        If (number) Then
            text1.Text = text1.Text & "3"
        Else
            text1.Text = "3"
            number = True
        End If
    End If
    change1 = True
End Sub

Private Sub number4_Click(Index As Integer)
    buf1 = text1.Text
    buf = CDbl(buf1)
    buf2 = buf
    point2 = Right(buf1, 1)
    If (buf1 <> buf2 And point2 <> ",") Then
        full = True
        text1.Text = Left(text1.Text, Len(text1.Text) - 1)
    End If
    If Not (full) Then
        If (number) Then
            text1.Text = text1.Text & "4"
        Else
            text1.Text = "4"
            number = True
        End If
    End If
    change1 = True
End Sub

Private Sub number5_Click(Index As Integer)
    buf1 = text1.Text
    buf = CDbl(buf1)
    buf2 = buf
    point2 = Right(buf1, 1)
    If (buf1 <> buf2 And point2 <> ",") Then
        full = True
        text1.Text = Left(text1.Text, Len(text1.Text) - 1)
    End If
    If Not (full) Then
        If (number) Then
            text1.Text = text1.Text & "5"
        Else
            text1.Text = "5"
            number = True
        End If
    End If
    change1 = True
End Sub

Private Sub number6_Click(Index As Integer)
    buf1 = text1.Text
    buf = CDbl(buf1)
    buf2 = buf
    point2 = Right(buf1, 1)
    If (buf1 <> buf2 And point2 <> ",") Then
        full = True
        text1.Text = Left(text1.Text, Len(text1.Text) - 1)
    End If
    If Not (full) Then
        If (number) Then
            text1.Text = text1.Text & "6"
        Else
            text1.Text = "6"
            number = True
        End If
    End If
    change1 = True
End Sub

Private Sub number7_Click(Index As Integer)
    buf1 = text1.Text
    buf = CDbl(buf1)
    buf2 = buf
    point2 = Right(buf1, 1)
    If (buf1 <> buf2 And point2 <> ",") Then
        full = True
        text1.Text = Left(text1.Text, Len(text1.Text) - 1)
    End If
    If Not (full) Then
        If (number) Then
            text1.Text = text1.Text & "7"
        Else
            text1.Text = "7"
            number = True
        End If
    End If
    change1 = True
End Sub

Private Sub number8_Click(Index As Integer)
    buf1 = text1.Text
    buf = CDbl(buf1)
    buf2 = buf
    point2 = Right(buf1, 1)
    If (buf1 <> buf2 And point2 <> ",") Then
        full = True
        text1.Text = Left(text1.Text, Len(text1.Text) - 1)
    End If
    If Not (full) Then
        If (number) Then
            text1.Text = text1.Text & "8"
        Else
            text1.Text = "8"
            number = True
        End If
    End If
    change1 = True
End Sub

Private Sub number9_Click(Index As Integer)
    Dim point2 As String
    buf1 = text1.Text
    buf = CDbl(buf1)
    buf2 = buf
    point2 = Right(buf1, 1)
    If (buf1 <> buf2 And point2 <> ",") Then
        full = True
        text1.Text = Left(text1.Text, Len(text1.Text) - 1)
    End If
    If Not (full) Then
        If (number) Then
            text1.Text = text1.Text & "9"
        Else
            text1.Text = "9"
            number = True
        End If
    End If
    change1 = True
End Sub

Private Sub point_Click(Index As Integer)
    If Not (point1) Then
        Dim new1 As Boolean
        new1 = (numder Or zero) And flag
        If flag And Not (number Or zero) Then
            text1.Text = "0,"
            point1 = True
            number = True
        Else
            text1.Text = text1.Text & ","
            point1 = True
            number = True
        End If
    End If
End Sub

Private Sub sign1_Click(Index As Integer)
    If (equal1 And (number Or zero)) Then
        a = text1.Text
        flag = False
    End If
    Dim press As Boolean
    press = zero Or number
    press = press And flag
    If (press) Then
        b = text1.Text
        If s = "+" Then
            text1.Text = a + b
        ElseIf s = "-" Then
            text1.Text = a - b
        ElseIf s = "*" Then
            text1.Text = a * b
        ElseIf s = "/" Then
            If b = 0 Then
                text1.Text = "You die, bustard!!"
                die = True
            Else
                text1.Text = a / b
            End If
        End If
    End If
    s = "+"
    flag = True
    number = False
    point1 = False
    zero = False
    change1 = False
    equal1 = True
    full = False
    If Not (die) Then
        a = text1.Text
    End If
End Sub

Private Sub sign2_Click(Index As Integer)
    Dim press As Boolean
    press = zero Or number
    press = press And flag
    If (press) Then
        b = text1.Text
        If s = "+" Then
            text1.Text = a + b
        ElseIf s = "-" Then
            text1.Text = a - b
        ElseIf s = "*" Then
            text1.Text = a * b
        ElseIf s = "/" Then
            If b = 0 Then
                text1.Text = "You die, bustard!!"
                die = True
            Else
                text1.Text = a / b
            End If
        End If
    End If
    s = "-"
    flag = True
    number = False
    point1 = False
    zero = False
    change1 = False
    equal1 = True
    full = False
    If Not (die) Then
        a = text1.Text
    End If
End Sub

Private Sub sign3_Click(Index As Integer)
    Dim press As Boolean
    press = zero Or number
    press = press And flag
    If (press) Then
        b = text1.Text
        If s = "+" Then
            text1.Text = a + b
        ElseIf s = "-" Then
            text1.Text = a - b
        ElseIf s = "*" Then
            text1.Text = a * b
        ElseIf s = "/" Then
            If b = 0 Then
                text1.Text = "You die, bustard!!"
                die = True
            Else
                text1.Text = a / b
            End If
        End If
    End If
    s = "*"
    flag = True
    number = False
    point1 = False
    zero = False
    change1 = False
    equal1 = True
    full = False
    If Not (die) Then
        a = text1.Text
    End If
End Sub

Private Sub sign4_Click(Index As Integer)
    Dim press As Boolean
    press = zero Or number
    press = press And flag
    If (press) Then
        b = text1.Text
        If s = "+" Then
            text1.Text = a + b
        ElseIf s = "-" Then
            text1.Text = a - b
        ElseIf s = "*" Then
            text1.Text = a * b
        ElseIf s = "/" Then
            If b = 0 Then
                text1.Text = "You die, bustard!!"
                die = True
            Else
                text1.Text = a / b
            End If
        End If
    End If
    s = "/"
    flag = True
    number = False
    point1 = False
    zero = False
    change1 = False
    equal1 = True
    full = False
    If Not (die) Then
        a = text1.Text
    End If
End Sub
