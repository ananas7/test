VERSION 5.00
Begin VB.Form Calc 
   BackColor       =   &H80000012&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calc"
   ClientHeight    =   4470
   ClientLeft      =   2865
   ClientTop       =   2175
   ClientWidth     =   4320
   FillColor       =   &H000000FF&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   4320
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
      Index           =   15
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
      Index           =   14
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
      Index           =   16
      Left            =   3480
      TabIndex        =   15
      Top             =   1920
      Width           =   735
   End
   Begin VB.CommandButton sign 
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
      Index           =   17
      Left            =   3480
      TabIndex        =   14
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton sign 
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
   Begin VB.CommandButton sign 
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
   Begin VB.CommandButton sign 
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
   Begin VB.CommandButton sign 
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
   Begin VB.CommandButton number 
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
   Begin VB.CommandButton number 
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
   Begin VB.CommandButton number 
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
   Begin VB.CommandButton number 
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
   Begin VB.CommandButton number 
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
   Begin VB.CommandButton number 
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
   Begin VB.CommandButton number 
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
   Begin VB.CommandButton number 
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
   Begin VB.CommandButton number 
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
   Begin VB.CommandButton number 
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
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
Dim press As Boolean
Dim buf1 As String
Dim buf As Double
Dim equal1 As Boolean
Dim die As Boolean
Dim press1 As Boolean
Dim number1 As Boolean


Private Sub change_Click(Index As Integer)
   If Not die Then
        buf1 = text1.Text
        buf = -CDbl(buf1)
        point2 = Right(buf1, 1)
        If point2 = "," Then
            text1.Text = buf & ","
        Else
            text1.Text = buf
        End If
        number1 = True
        If equal1 Then
            a = text1.Text
        End If
    End If
End Sub

Private Sub delete2_Click(Index As Integer)
    text1.Text = "0"
    s = ""
    die = False
    press = False
    equal1 = False
    number1 = False
    a = 0
    b = 0
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    text1.SetFocus
    Debug.Print KeyAscii
    If (Chr(KeyAscii) >= "0" And Chr(KeyAscii) <= "9") Then
        number_Click (Int(Chr(KeyAscii)))
    ElseIf (KeyAscii = 8 Or KeyAscii = 46) Then
        delete2_Click (0)
    ElseIf (Chr(KeyAscii) = "+") Then
        sign_Click (10)
    ElseIf (Chr(KeyAscii) = "-") Then
        sign_Click (11)
    ElseIf (Chr(KeyAscii) = "*") Then
        sign_Click (12)
    ElseIf (Chr(KeyAscii) = "/") Then
        sign_Click (13)
    ElseIf (KeyAscii = 13 Or Chr(KeyAscii) = "=") Then
        sign_Click (17)
    ElseIf Chr(KeyAscii) = "," Then
        point_Click (0)
    End If
    KeyAscii = 0
End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Debug.Print KeyCode
    If KeyCode = 46 Then
        delete2_Click (0)
    ElseIf KeyCode = 120 Then
        change_Click (0)
    End If
    KeyCode = 0
End Sub

Private Sub number_Click(Index As Integer)
    If Not die Then
        If press And Not equal1 Then
            text1.Text = "0"
            press = False
        End If
        If text1.Text = "0" Then
            text1.Text = ""
        End If
        buf1 = text1.Text & number(Index).Caption
        If buf1 = CStr(CDbl(buf1)) Then
            text1.Text = text1.Text & number(Index).Caption
        End If
        number1 = True
    End If
End Sub


Private Sub point_Click(Index As Integer)
    buf1 = text1.Text
    If equal1 Then
        text1.Text = "0"
    End If
    If InStr(1, text1.Text, ",") = 0 And Not die Then
        text1.Text = text1.Text & ","
    End If
End Sub

Private Sub sign_Click(Index As Integer)
    If Not die Then
        If s = "" Then
            a = text1.Text
        ElseIf Not equal1 And number1 Then
            b = text1.Text
        End If
        If Index = 10 Then
            s = "+"
            If equal1 Then
                equal1 = False
                a = text1.Text
                b = 0
            End If
        ElseIf Index = 11 Then
            s = "-"
            If equal1 Then
                equal1 = False
                a = text1.Text
                b = 0
            End If
        ElseIf Index = 12 Then
            s = "*"
            If equal1 Then
                equal1 = False
                a = text1.Text
                b = 0
            End If
        ElseIf Index = 13 Then
            s = "/"
            If equal1 Then
                equal1 = False
                a = text1.Text
                b = 1
            End If
        ElseIf Index = 17 Then
            equal1 = True
        End If
        If number1 Then
            If s = "+" Then
                text1.Text = a + b
            ElseIf s = "-" Then
                text1.Text = a - b
            ElseIf s = "*" Then
                text1.Text = a * b
            ElseIf s = "/" Then
                If b <> 0 Then
                    text1.Text = a / b
                Else
                    text1.Text = "YOU DIE BUSTERD!!!"
                    die = True
                End If
            End If
            If Not die Then
                a = text1.Text
            End If
            number1 = False
        End If
        press = True
    End If
End Sub
