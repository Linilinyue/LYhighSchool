VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Tanker   202小越子"
   ClientHeight    =   9330
   ClientLeft      =   2385
   ClientTop       =   1305
   ClientWidth     =   15690
   LinkTopic       =   "Form1"
   ScaleHeight     =   9330
   ScaleWidth      =   15690
   Begin VB.Timer Timer1 
      Interval        =   30
      Left            =   15120
      Top             =   480
   End
   Begin VB.CommandButton Command1 
      Caption         =   "开始"
      Height          =   495
      Left            =   14040
      TabIndex        =   11
      Top             =   8760
      Width           =   1095
   End
   Begin VB.Shape Shape6 
      Height          =   135
      Index           =   3
      Left            =   12840
      Top             =   4680
      Width           =   135
   End
   Begin VB.Shape Shape6 
      Height          =   135
      Index           =   2
      Left            =   12120
      Top             =   4680
      Width           =   135
   End
   Begin VB.Shape Shape6 
      Height          =   135
      Index           =   1
      Left            =   12480
      Top             =   5040
      Width           =   135
   End
   Begin VB.Shape Shape6 
      Height          =   135
      Index           =   0
      Left            =   12480
      Top             =   4320
      Width           =   135
   End
   Begin VB.Shape Shape5 
      Height          =   135
      Index           =   3
      Left            =   720
      Top             =   4680
      Width           =   135
   End
   Begin VB.Shape Shape5 
      Height          =   135
      Index           =   2
      Left            =   0
      Top             =   4680
      Width           =   135
   End
   Begin VB.Shape Shape5 
      Height          =   135
      Index           =   1
      Left            =   360
      Top             =   5040
      Width           =   135
   End
   Begin VB.Shape Shape5 
      Height          =   135
      Index           =   0
      Left            =   360
      Top             =   4320
      Width           =   135
   End
   Begin VB.Label Label5 
      Caption         =   "0"
      Height          =   255
      Index           =   3
      Left            =   15240
      TabIndex        =   17
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Label5 
      Caption         =   "0"
      Height          =   255
      Index           =   2
      Left            =   14640
      TabIndex        =   16
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Label5 
      Caption         =   "0"
      Height          =   255
      Index           =   1
      Left            =   13800
      TabIndex        =   15
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Label5 
      Caption         =   "0"
      Height          =   255
      Index           =   0
      Left            =   13320
      TabIndex        =   14
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Label4 
      Caption         =   "右方KD"
      Height          =   495
      Left            =   14640
      TabIndex        =   13
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "左方KD"
      Height          =   495
      Left            =   13320
      TabIndex        =   12
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "场外"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   14160
      TabIndex        =   10
      Top             =   360
      Width           =   735
   End
   Begin VB.Shape Shape4 
      Height          =   615
      Index           =   1
      Left            =   12240
      Top             =   4440
      Width           =   615
   End
   Begin VB.Shape Shape4 
      Height          =   615
      Index           =   0
      Left            =   120
      Top             =   4440
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "20"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   9
      Left            =   12360
      TabIndex        =   9
      Top             =   4560
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "20"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   8
      Left            =   240
      TabIndex        =   8
      Top             =   4560
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   7
      Left            =   12240
      TabIndex        =   7
      Top             =   8640
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   6
      Left            =   12240
      TabIndex        =   6
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   5
      Left            =   480
      TabIndex        =   5
      Top             =   8760
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   4
      Left            =   360
      TabIndex        =   4
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "50"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   3
      Left            =   9000
      TabIndex        =   3
      Top             =   6840
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "50"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   2
      Left            =   9120
      TabIndex        =   2
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "50"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   1
      Left            =   4080
      TabIndex        =   1
      Top             =   6840
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "50"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   0
      Left            =   4080
      TabIndex        =   0
      Top             =   2280
      Width           =   375
   End
   Begin VB.Line Line1 
      X1              =   13200
      X2              =   13200
      Y1              =   -120
      Y2              =   9240
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      Height          =   855
      Index           =   3
      Left            =   8760
      Top             =   6600
      Width           =   855
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      Height          =   855
      Index           =   2
      Left            =   8880
      Top             =   2040
      Width           =   855
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      Height          =   855
      Index           =   1
      Left            =   3840
      Top             =   6600
      Width           =   855
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      Height          =   855
      Index           =   0
      Left            =   3840
      Top             =   2040
      Width           =   855
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0C0FF&
      BackStyle       =   1  'Opaque
      Height          =   1815
      Index           =   1
      Left            =   11400
      Top             =   3840
      Width           =   1815
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0C0FF&
      BackStyle       =   1  'Opaque
      Height          =   1815
      Index           =   0
      Left            =   0
      Top             =   3840
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   1  'Opaque
      Height          =   1695
      Index           =   3
      Left            =   11880
      Shape           =   2  'Oval
      Top             =   7920
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   1  'Opaque
      Height          =   1695
      Index           =   2
      Left            =   11760
      Shape           =   2  'Oval
      Top             =   -240
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   1  'Opaque
      Height          =   1695
      Index           =   1
      Left            =   -240
      Shape           =   2  'Oval
      Top             =   8040
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   1  'Opaque
      Height          =   1695
      Index           =   0
      Left            =   -240
      Shape           =   2  'Oval
      Top             =   -360
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim u1 As Boolean
Dim u2 As Boolean
Dim d1 As Boolean
Dim d2 As Boolean
Dim l1 As Boolean
Dim l2 As Boolean
Dim r1 As Boolean
Dim r2 As Boolean
Dim pao As Integer

Private Sub command1_keydown(keycode As Integer, shift As Integer)
    If keycode = vbKeyW Then u1 = True
    If keycode = 104 Then u2 = True
    If keycode = vbKeyS Then d1 = True
    If keycode = 101 Then d2 = True
    If keycode = vbKeyA Then l1 = True
    If keycode = 100 Then l2 = True
    If keycode = vbKeyD Then r1 = True
    If keycode = 102 Then r2 = True
End Sub

Private Sub command1_keyup(keycode As Integer, shift As Integer)
    If keycode = vbKeyW Then u1 = False
    If keycode = 104 Then u2 = False
    If keycode = vbKeyS Then d1 = False
    If keycode = 101 Then d2 = False
    If keycode = vbKeyA Then l1 = False
    If keycode = 100 Then l2 = Falses
    If keycode = vbKeyD Then r1 = False
    If keycode = 102 Then r2 = False
End Sub

Private Sub Form_Load()
Me.KeyPreview = True
    u1 = False
    u2 = False
    d1 = False
    d2 = False
    l1 = False
    l2 = False
    r1 = False
    r2 = False
    For pao = 0 To 3
        Shape5(pao).Visible = False
        Shape6(pao).Visible = False
    Next pao
End Sub

Private Sub Timer1_Timer()

'这个timer用来实现同时响应多个按键
    
    '左方的移动
    If u1 = True Then
        Shape5(0).Visible = True
        Shape5(1).Visible = False
        Shape5(2).Visible = False
        Shape5(3).Visible = False
        If Shape4(0).Top >= 0 Then '上
            Shape4(0).Top = Shape4(0).Top - 100
            Label1(8).Top = Label1(8).Top - 100
            For pao = 0 To 3
                Shape5(pao).Top = Shape5(pao).Top - 100
            Next pao
        End If
    End If
    
    If d1 = True Then
        Shape5(0).Visible = False
        Shape5(1).Visible = True
        Shape5(2).Visible = False
        Shape5(3).Visible = False
        If Shape4(0).Top <= Form1.Height - Shape4(0).Height * 2 Then
            Shape4(0).Top = Shape4(0).Top + 100
            Label1(8).Top = Label1(8).Top + 100
            For pao = 0 To 3
                Shape5(pao).Top = Shape5(pao).Top + 100
            Next pao
        End If
    End If
    
    
    If l1 = True Then
        Shape5(0).Visible = False
        Shape5(1).Visible = False
        Shape5(2).Visible = True
        Shape5(3).Visible = False
        If Shape4(0).Left >= 0 Then
            Shape4(0).Left = Shape4(0).Left - 100
            Label1(8).Left = Label1(8).Left - 100
            For pao = 0 To 3
                Shape5(pao).Left = Shape5(pao).Left - 100
            Next pao
        End If
    End If
    
    
    If r1 = True Then
        Shape5(0).Visible = False
        Shape5(1).Visible = False
        Shape5(2).Visible = False
        Shape5(3).Visible = True
        If Shape4(0).Left <= 13200 - 700 Then
            Shape4(0).Left = Shape4(0).Left + 100
            Label1(8).Left = Label1(8).Left + 100
            For pao = 0 To 3
                Shape5(pao).Left = Shape5(pao).Left + 100
            Next pao
        End If
    End If
    
    '右方的移动
    
    If u2 = True Then
        Shape6(0).Visible = True
        Shape6(1).Visible = False
        Shape6(2).Visible = False
        Shape6(3).Visible = False
        If Shape4(1).Top >= 0 Then
            Shape4(1).Top = Shape4(1).Top - 100
            Label1(9).Top = Label1(9).Top - 100
            For pao = 0 To 3
                Shape6(pao).Top = Shape6(pao).Top - 100
            Next pao
        End If
    End If
    
    If d2 = True Then
        Shape6(0).Visible = False
        Shape6(1).Visible = True
        Shape6(2).Visible = False
        Shape6(3).Visible = False
        If Shape4(1).Top <= Form1.Height - Shape4(1).Height * 2 Then
            Shape4(1).Top = Shape4(1).Top + 100
            Label1(9).Top = Label1(9).Top + 100
            For pao = 0 To 3
                Shape6(pao).Top = Shape6(pao).Top + 100
            Next pao
        End If
    End If
    
    If l2 = True Then
        Shape6(0).Visible = False
        Shape6(1).Visible = False
        Shape6(2).Visible = True
        Shape6(3).Visible = False
        If Shape4(1).Left >= 0 Then
            Shape4(1).Left = Shape4(1).Left - 100
            Label1(9).Left = Label1(9).Left - 100
            For pao = 0 To 3
                Shape6(pao).Left = Shape6(pao).Left - 100
            Next pao
        End If
    End If
    
    If r2 = True Then
        Shape6(0).Visible = False
        Shape6(1).Visible = False
        Shape6(2).Visible = False
        Shape6(3).Visible = True
        If Shape4(1).Left <= 13200 - 600 Then
            Shape4(1).Left = Shape4(1).Left + 100
            Label1(9).Left = Label1(9).Left + 100
            For pao = 0 To 3
                Shape6(pao).Left = Shape6(pao).Left + 100
            Next pao
        End If
    End If
End Sub
