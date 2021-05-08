VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Tanker   202小越子"
   ClientHeight    =   9255
   ClientLeft      =   2385
   ClientTop       =   1305
   ClientWidth     =   15690
   LinkTopic       =   "Form1"
   ScaleHeight     =   9255
   ScaleWidth      =   15690
   Begin VB.Timer Timer5 
      Interval        =   10
      Left            =   15120
      Top             =   6120
   End
   Begin VB.Timer Timer4 
      Interval        =   10
      Left            =   15120
      Top             =   7320
   End
   Begin VB.Timer Timer3 
      Interval        =   10
      Left            =   15120
      Top             =   6720
   End
   Begin VB.Timer Timer2 
      Interval        =   10
      Left            =   14640
      Top             =   6120
   End
   Begin VB.Timer Timer1 
      Interval        =   15
      Left            =   15120
      Top             =   5520
   End
   Begin VB.CommandButton Command1 
      Caption         =   "开始"
      Height          =   495
      Left            =   14040
      TabIndex        =   11
      Top             =   8760
      Width           =   1095
   End
   Begin VB.Line Line6 
      X1              =   14880
      X2              =   15240
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line Line5 
      X1              =   15240
      X2              =   15240
      Y1              =   2640
      Y2              =   4440
   End
   Begin VB.Line Line4 
      X1              =   14880
      X2              =   15240
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line3 
      X1              =   14880
      X2              =   14880
      Y1              =   2640
      Y2              =   4440
   End
   Begin VB.Shape Shape8 
      Height          =   135
      Index           =   6
      Left            =   15000
      Top             =   4200
      Width           =   135
   End
   Begin VB.Shape Shape8 
      Height          =   135
      Index           =   5
      Left            =   15000
      Top             =   3960
      Width           =   135
   End
   Begin VB.Shape Shape8 
      Height          =   135
      Index           =   4
      Left            =   15000
      Top             =   3720
      Width           =   135
   End
   Begin VB.Shape Shape8 
      Height          =   135
      Index           =   3
      Left            =   15000
      Top             =   3480
      Width           =   135
   End
   Begin VB.Shape Shape8 
      Height          =   135
      Index           =   2
      Left            =   15000
      Top             =   3240
      Width           =   135
   End
   Begin VB.Shape Shape8 
      Height          =   135
      Index           =   1
      Left            =   15000
      Top             =   3120
      Width           =   135
   End
   Begin VB.Shape Shape8 
      Height          =   135
      Index           =   0
      Left            =   15000
      Top             =   2760
      Width           =   135
   End
   Begin VB.Line Line8 
      X1              =   13560
      X2              =   13920
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line11 
      X1              =   13920
      X2              =   13920
      Y1              =   2640
      Y2              =   4440
   End
   Begin VB.Line Line10 
      X1              =   13560
      X2              =   13560
      Y1              =   2640
      Y2              =   4440
   End
   Begin VB.Line Line9 
      X1              =   13560
      X2              =   13920
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Shape Shape4 
      Height          =   1335
      Index           =   9
      Left            =   11880
      Top             =   7920
      Width           =   1335
   End
   Begin VB.Shape Shape4 
      Height          =   1335
      Index           =   8
      Left            =   0
      Top             =   7920
      Width           =   1335
   End
   Begin VB.Shape Shape4 
      Height          =   1335
      Index           =   7
      Left            =   11880
      Top             =   0
      Width           =   1335
   End
   Begin VB.Shape Shape4 
      Height          =   1335
      Index           =   6
      Left            =   0
      Top             =   0
      Width           =   1335
   End
   Begin VB.Shape Shape4 
      Height          =   615
      Index           =   5
      Left            =   9000
      Top             =   6600
      Width           =   615
   End
   Begin VB.Shape Shape4 
      Height          =   615
      Index           =   4
      Left            =   3960
      Top             =   6600
      Width           =   615
   End
   Begin VB.Shape Shape4 
      Height          =   615
      Index           =   3
      Left            =   9000
      Top             =   2160
      Width           =   615
   End
   Begin VB.Shape Shape4 
      Height          =   615
      Index           =   2
      Left            =   3960
      Top             =   2160
      Width           =   615
   End
   Begin VB.Line Line2 
      X1              =   13200
      X2              =   120
      Y1              =   9240
      Y2              =   9240
   End
   Begin VB.Shape Shape7 
      Height          =   135
      Index           =   6
      Left            =   13680
      Top             =   4200
      Width           =   135
   End
   Begin VB.Shape Shape7 
      Height          =   135
      Index           =   5
      Left            =   13680
      Top             =   3960
      Width           =   135
   End
   Begin VB.Shape Shape7 
      Height          =   135
      Index           =   4
      Left            =   13680
      Top             =   3720
      Width           =   135
   End
   Begin VB.Shape Shape7 
      Height          =   135
      Index           =   3
      Left            =   13680
      Top             =   3480
      Width           =   135
   End
   Begin VB.Shape Shape7 
      Height          =   135
      Index           =   2
      Left            =   13680
      Top             =   3240
      Width           =   135
   End
   Begin VB.Shape Shape7 
      Height          =   135
      Index           =   1
      Left            =   13680
      Top             =   3000
      Width           =   135
   End
   Begin VB.Shape Shape7 
      Height          =   135
      Index           =   0
      Left            =   13680
      Top             =   2760
      Width           =   135
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
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   14640
      TabIndex        =   13
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "左方KD"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
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
      Index           =   9
      Left            =   12240
      TabIndex        =   9
      Top             =   8400
      Width           =   735
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
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   8
      Left            =   480
      TabIndex        =   8
      Top             =   8280
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
      Index           =   7
      Left            =   12240
      TabIndex        =   7
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
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   6
      Left            =   360
      TabIndex        =   6
      Top             =   480
      Width           =   615
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
      Height          =   615
      Index           =   5
      Left            =   9120
      TabIndex        =   5
      Top             =   6720
      Width           =   855
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
      ForeColor       =   &H00FF0000&
      Height          =   495
      Index           =   4
      Left            =   4080
      TabIndex        =   4
      Top             =   6720
      Width           =   735
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
      Height          =   615
      Index           =   3
      Left            =   9120
      TabIndex        =   3
      Top             =   2280
      Width           =   615
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
      ForeColor       =   &H00FF0000&
      Height          =   735
      Index           =   2
      Left            =   4080
      TabIndex        =   2
      Top             =   2280
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
      Index           =   1
      Left            =   12240
      TabIndex        =   1
      Top             =   4560
      Width           =   1815
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
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   4560
      Width           =   2055
   End
   Begin VB.Line Line1 
      X1              =   13200
      X2              =   13200
      Y1              =   -120
      Y2              =   9240
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
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'说明
Dim sm As Integer

'运动部分
Dim u1 As Boolean
Dim u2 As Boolean
Dim d1 As Boolean
Dim d2 As Boolean
Dim l1 As Boolean
Dim l2 As Boolean
Dim r1 As Boolean
Dim r2 As Boolean
Dim pao As Integer

'导弹部分
Dim i As Integer
Dim m As Integer
Dim a(0 To 6) As Boolean        '运动，判断是否到底
Dim a2(0 To 6) As Boolean
Dim b As Boolean                '每次发一个
Dim b2 As Boolean
Dim c(0 To 6) As Boolean        '从tanker位置发射
Dim c2(0 To 6) As Boolean
Dim dd1 As Boolean
Dim dd2 As Boolean
Dim fdd1(0 To 6) As Integer      '让导弹一次到底
Dim fdd2(0 To 6) As Integer
Dim j As Integer
Dim n As Integer

'伤害部分
Dim hurt(0 To 9) As Boolean
Dim h As Integer
Dim h2 As Integer
Dim zsj As Boolean
Dim zxj As Boolean
Dim ysj As Boolean
Dim yxj As Boolean

'胜利部分
Dim zsl As Integer
Dim ysl As Integer


Private Sub command1_keydown(keycode As Integer, shift As Integer)
    '运动部分
    If keycode = vbKeyW Then
        u1 = True
         For i = 0 To 6
            If fdd1(i) = 0 Then
                fdd1(i) = 1
                Exit For
            End If
        Next i
    End If
    
    If keycode = 104 Then
        u2 = True
         For i = 0 To 6
            If fdd2(i) = 0 Then
                fdd2(i) = 1
                Exit For
            End If
        Next i
    End If

    
    If keycode = vbKeyS Then
        d1 = True
         For i = 0 To 6
            If fdd1(i) = 0 Then
                fdd1(i) = 2
                Exit For
            End If
        Next i
    End If
    
    If keycode = 101 Then
        d2 = True
         For i = 0 To 6
            If fdd2(i) = 0 Then
                fdd2(i) = 2
                Exit For
            End If
        Next i
    End If
        
    If keycode = vbKeyA Then
        l1 = True
        For i = 0 To 6
            If fdd1(i) = 0 Then
                fdd1(i) = 3
                Exit For
            End If
        Next i
    End If
    
    
    If keycode = 100 Then
        l2 = True
         For i = 0 To 6
            If fdd2(i) = 0 Then
                fdd2(i) = 3
                Exit For
            End If
        Next i
    End If
    
    If keycode = vbKeyD Then
        r1 = True
        For i = 0 To 6
            If fdd1(i) = 0 Then
                fdd1(i) = 4
                Exit For
            End If
        Next i
    End If
    
    If keycode = 102 Then
        r2 = True
         For i = 0 To 6
            If fdd2(i) = 0 Then
                fdd2(i) = 4
                Exit For
            End If
        Next i
    End If
    
    
    '导弹部分
        '左方
        If keycode = vbKeyJ Then
        b = True
            For i = 0 To 6
                If a(i) = False And b = True Then
                    a(i) = True
                    b = False
                    c(i) = True
                End If
            Next i
        End If
        
        '右方
        If keycode = 107 Then
        b2 = True
            For m = 0 To 6
                If a2(m) = False And b2 = True Then
                    a2(m) = True
                    b2 = False
                    c2(m) = True
                End If
            Next m
        End If
    
End Sub

Private Sub command1_keyup(keycode As Integer, shift As Integer)
    If keycode = vbKeyW Then
        u1 = False
    End If
    
    If keycode = 104 Then u2 = False
    
    If keycode = vbKeyS Then
        d1 = False
    End If
    
    If keycode = 101 Then d2 = False
    
    If keycode = vbKeyA Then
        l1 = False
    End If
    
    If keycode = 100 Then l2 = False
    
    If keycode = vbKeyD Then
        r1 = False
    End If
    
    If keycode = 102 Then r2 = False
    
    '导弹部分
    If keycode = vbKeyJ Then
        b = False
    End If
    
    If keycode = 107 Then
        b2 = False
    End If
    
End Sub

Private Sub Form_Load()
    sm = MsgBox("欢迎来到Tanker", 0, "欢迎")
    sm = MsgBox("左方移动为WASD，J键发炮；右方移动为小键盘8456，6右边的+号发炮", 0, "说明")
    sm = MsgBox("目的是击败对方边角两个大方形基地。要先击败对方小防御塔。攻击对方防御塔会扣自己血。允许攻击自己防御塔。", 0, "说明")
    sm = MsgBox("血量扣光则自动回到基地，并瞬间恢复至5血", 0, "说明")
    sm = MsgBox("就这么多，请注意切换输入法和打开小键盘锁。祝你游戏愉快", 0, "点击确定即开始")
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
    For i = 0 To 6
        a(i) = False
        a2(i) = False
        fdd1(i) = 4
        fdd2(i) = 4
    Next i
    
    For i = 0 To 9
        hurt(i) = True
    Next i
    zsj = False
    zxj = False
    ysj = False
    yxj = False
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
            Label1(0).Top = Label1(0).Top - 100
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
            Label1(0).Top = Label1(0).Top + 100
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
        If Shape4(0).Left > 0 Then
            Shape4(0).Left = Shape4(0).Left - 100
            Label1(0).Left = Label1(0).Left - 100
            If Shape4(0).Left < 0 Then
                Shape4(0).Left = 0
                Label1(0).Left = Label1(0).Left + 100
                For pao = 0 To 3
                    Shape5(pao).Left = Shape5(pao).Left + 80
                Next pao
            End If
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
            Label1(0).Left = Label1(0).Left + 100
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
            Label1(1).Top = Label1(1).Top - 100
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
        If Shape4(1).Top < Form1.Height - Shape4(1).Height * 2 Then
            Shape4(1).Top = Shape4(1).Top + 100
            Label1(1).Top = Label1(1).Top + 100
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
            Label1(1).Left = Label1(1).Left - 100
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
        If Shape4(1).Left < 13200 - 600 Then
            Shape4(1).Left = Shape4(1).Left + 100
            Label1(1).Left = Label1(1).Left + 100
            If Shape4(1).Left > 13200 - 600 Then
                Shape4(1).Left = 13200 - 600
                For pao = 0 To 3
                    Shape6(pao).Left = Shape6(pao).Left - 45
                Next pao
            End If
            For pao = 0 To 3
                Shape6(pao).Left = Shape6(pao).Left + 100
            Next pao
        End If
    End If
    
    
End Sub

Private Sub Timer2_Timer()
'导弹部分


'左方
If Shape5(2).Visible = True Then
    For i = 0 To 6
        If a(i) = True Then
            If c(i) = True Then
                Shape7(i).Left = Shape5(2).Left
                Shape7(i).Top = Shape5(2).Top
                c(i) = False
            End If
            j = i
            If fdd1(j) = 4 Then fdd1(j) = 2
        End If
    Next i
ElseIf Shape5(0).Visible = True Then
    For i = 0 To 6
        If a(i) = True Then
            If c(i) = True Then
                Shape7(i).Left = Shape5(0).Left
                Shape7(i).Top = Shape5(0).Top
                c(i) = False
            End If
            j = i
            If fdd1(j) = 4 Then fdd1(j) = 5
        End If
    Next i
ElseIf Shape5(1).Visible = True Then
    For i = 0 To 6
        If a(i) = True Then
            If c(i) = True Then
                Shape7(i).Left = Shape5(1).Left
                Shape7(i).Top = Shape5(1).Top
                c(i) = False
            End If
            j = i
            If fdd1(j) = 4 Then fdd1(j) = 1
        End If
    Next i
ElseIf Shape5(3).Visible = True Then
    For i = 0 To 6
        If a(i) = True Then
            If c(i) = True Then
                Shape7(i).Left = Shape5(3).Left
                Shape7(i).Top = Shape5(3).Top
                c(i) = False
            End If
            j = i
            If fdd1(j) = 4 Then fdd1(j) = 3
        End If
    Next i
End If





End Sub

Private Sub Timer3_Timer()
    '伤害判定
    
    '回血
    If 3840 < Shape4(0).Top And Shape4(0).Top < 5640 And Shape4(0).Left < 13200 And 11400 < Shape4(0).Left And Val(Label1(0).Caption) <= 20 Then
        Label1(0).Caption = Str(Val(Label1(0).Caption) - 0.1)
    ElseIf 3840 < Shape4(0).Top And Shape4(0).Top < 5640 And Shape4(0).Left < 1815 And 0 <= Shape4(0).Left And Val(Label1(0).Caption) < 20 Then
        Label1(0).Caption = Str(Val(Label1(0).Caption) + 0.1)
    End If
    
    If 3840 < Shape4(1).Top And Shape4(1).Top < 5640 And Shape4(1).Left < 13200 And 11400 < Shape4(1).Left And Val(Label1(1).Caption) < 20 Then
        Label1(1).Caption = Str(Val(Label1(1).Caption) + 0.1)
    ElseIf 3840 < Shape4(1).Top And Shape4(1).Top < 5640 And Shape4(1).Left < 1815 And 0 <= Shape4(1).Left And Val(Label1(1).Caption) <= 20 Then
        Label1(1).Caption = Str(Val(Label1(1).Caption) - 0.1)
    End If
    
    '左方 伤害
    For i = 0 To 9
        If hurt(i) = True Then
            For h = 0 To 6
                If Shape7(h).Top <= 1320 And Shape7(h).Left <= 1320 And zsj = True Then
                    Label1(6).Caption = Str(Val(Label1(6).Caption) - 1)
                    Shape7(h).Left = -1000
                    Shape7(h).Top = -1000
                ElseIf Shape7(h).Top <= 1320 And Shape7(h).Left >= 11880 And ysj = True Then
                    Label1(7).Caption = Str(Val(Label1(7).Caption) - 1)
                    Shape7(h).Left = -1000
                    Shape7(h).Top = -1000
                ElseIf Shape7(h).Top >= 7920 And Shape7(h).Left <= 1320 And zxj = True Then
                    Label1(8).Caption = Str(Val(Label1(8).Caption) - 1)
                    Shape7(h).Left = -1000
                    Shape7(h).Top = -1000
                ElseIf Shape7(h).Top >= 7920 And Shape7(h).Left >= 11880 And yxj = True Then
                    Label1(9).Caption = Str(Val(Label1(9).Caption) - 1)
                    Shape7(h).Left = -1000
                    Shape7(h).Top = -1000
                ElseIf 2160 <= Shape7(h).Top And Shape7(h).Top <= 2875 And 3960 <= Shape7(h).Left And Shape7(h).Left <= 4575 And Val(Label1(2).Caption) > 0 Then
                    Label1(2).Caption = Str(Val(Label1(2).Caption) - 1)
                    If Label1(2).Caption = 0 Then zsj = True
                    Shape7(h).Left = -1000
                    Shape7(h).Top = -1000
                ElseIf 2160 <= Shape7(h).Top And Shape7(h).Top <= 2875 And 9000 <= Shape7(h).Left And Shape7(h).Left <= 9615 And Val(Label1(3).Caption) > 0 Then
                    Label1(3).Caption = Str(Val(Label1(3).Caption) - 1)
                    Label1(0).Caption = Str(Val(Label1(0).Caption) - 1)
                    If Label1(3).Caption = 0 Then ysj = True
                    If Label1(3).Caption = 0 Then Shape7(h).Left = -1000
                    Shape7(h).Top = -1000
                ElseIf 6600 <= Shape7(h).Top And Shape7(h).Top <= 7215 And 3960 <= Shape7(h).Left And Shape7(h).Left <= 4575 And Val(Label1(4).Caption) > 0 Then
                    Label1(4).Caption = Str(Val(Label1(4).Caption) - 1)
                    If Label1(4).Caption = 0 Then zxj = True
                    Shape7(h).Left = -1000
                    Shape7(h).Top = -1000
                ElseIf 6600 <= Shape7(h).Top And Shape7(h).Top <= 7215 And 9000 <= Shape7(h).Left And Shape7(h).Left <= 9615 And Val(Label1(5).Caption) > 0 Then
                    Label1(5).Caption = Str(Val(Label1(5).Caption) - 1)
                    Label1(0).Caption = Str(Val(Label1(0).Caption) - 1)
                    If Label1(5).Caption = 0 Then yxj = True
                    Shape7(h).Left = -1000
                    Shape7(h).Top = -1000
                ElseIf Shape4(1).Top <= Shape7(h).Top And Shape7(h).Top <= Shape4(1).Top + 615 And Shape4(1).Left <= Shape7(h).Left And Shape7(h).Left <= Shape4(1).Left + 615 Then
                    Label1(1).Caption = Str(Val(Label1(1).Caption) - 1)
                    Shape7(h).Left = -1000
                    Shape7(h).Top = -1000
                End If
                
                If Shape7(h).Top < 0 Or Shape7(h).Top > 9240 Or Shape7(h).Left < 0 Or Shape7(h).Left > 13680 Then
                    Shape7(h).Left = 13680
                    Shape7(h).Top = 2760 + 240 * h
                    a(h) = False
                    fdd1(h) = 4
                End If
            Next h
        End If
    Next i
    
    '右方 伤害
    
        For m = 0 To 9
        If hurt(m) = True Then
            For h2 = 0 To 6
                If Shape8(h2).Top <= 1320 And Shape8(h2).Left <= 1320 And zsj = True Then
                    Label1(6).Caption = Str(Val(Label1(6).Caption) - 1)
                    Shape8(h2).Left = -1000
                    Shape8(h2).Top = -1000
                ElseIf Shape8(h2).Top <= 1320 And Shape8(h2).Left >= 11880 And ysj = True Then
                    Label1(8).Caption = Str(Val(Label1(7).Caption) - 1)
                    Shape8(h2).Left = -1000
                    Shape8(h2).Top = -1000
                ElseIf Shape8(h2).Top >= 7920 And Shape8(h2).Left <= 1320 And zxj = True Then
                    Label1(8).Caption = Str(Val(Label1(8).Caption) - 1)
                    Shape8(h2).Left = -1000
                    Shape8(h2).Top = -1000
                ElseIf Shape8(h2).Top >= 7920 And Shape8(h2).Left >= 11880 And yxj = True Then
                    Label1(9).Caption = Str(Val(Label1(9).Caption) - 1)
                    Shape8(h2).Left = -1000
                    Shape8(h2).Top = -1000
                ElseIf 2160 <= Shape8(h2).Top And Shape8(h2).Top <= 2875 And 3960 <= Shape8(h2).Left And Shape8(h2).Left <= 4575 And Val(Label1(2).Caption) > 0 Then
                    Label1(2).Caption = Str(Val(Label1(2).Caption) - 1)
                    Label1(1).Caption = Str(Val(Label1(1).Caption) - 1)
                    If Label1(2).Caption = 0 Then zsj = True
                    Shape8(h2).Left = -1000
                    Shape8(h2).Top = -1000
                ElseIf 2160 <= Shape8(h2).Top And Shape8(h2).Top <= 2875 And 9000 <= Shape8(h2).Left And Shape8(h2).Left <= 9615 And Val(Label1(3).Caption) > 0 Then
                    Label1(3).Caption = Str(Val(Label1(3).Caption) - 1)
                    If Label1(3).Caption = 0 Then ysj = True
                    If Label1(3).Caption = 0 Then Shape8(h2).Left = -1000
                    Shape8(h2).Top = -1000
                ElseIf 6600 <= Shape8(h2).Top And Shape8(h2).Top <= 7215 And 3960 <= Shape8(h2).Left And Shape8(h2).Left <= 4575 And Val(Label1(4).Caption) > 0 Then
                    Label1(4).Caption = Str(Val(Label1(4).Caption) - 1)
                    Label1(1).Caption = Str(Val(Label1(1).Caption) - 1)
                    If Label1(4).Caption = 0 Then zxj = True
                    Shape8(h2).Left = -1000
                    Shape7(h2).Top = -1000
                ElseIf 6600 <= Shape8(h2).Top And Shape8(h2).Top <= 7215 And 9000 <= Shape8(h2).Left And Shape8(h2).Left <= 9615 And Val(Label1(5).Caption) > 0 Then
                    Label1(5).Caption = Str(Val(Label1(5).Caption) - 1)
                    If Label1(5).Caption = 0 Then yxj = True
                    Shape8(h2).Left = -1000
                    Shape8(h2).Top = -1000
                ElseIf Shape4(0).Top <= Shape8(h2).Top And Shape8(h2).Top <= Shape4(0).Top + 615 And Shape4(0).Left <= Shape8(h2).Left And Shape8(h2).Left <= Shape4(0).Left + 615 Then
                    Label1(0).Caption = Str(Val(Label1(0).Caption) - 1)
                    Shape8(h2).Left = -1000
                    Shape8(h2).Top = -1000
                End If
                
                If Shape8(h2).Top < 0 Or Shape8(h2).Top > 9240 Or Shape8(h2).Left < 0 Or Shape8(h2).Left > 13680 Then
                    Shape8(h2).Left = 15000
                    Shape8(h2).Top = 2760 + 240 * h2
                    a2(h2) = False
                    fdd2(h2) = 4
                End If
            Next h2
        End If
    Next m
    
    '生死判定
    If Val(Label1(0).Caption) <= 0 Then
        Label5(1).Caption = Str(Val(Label5(1).Caption) + 1)
        Label5(2).Caption = Str(Val(Label5(2).Caption) + 1)
        Label1(0).Caption = 5
        Shape4(0).Top = 4440
        Shape4(0).Left = 120
        Shape5(0).Left = 360
        Shape5(0).Top = 4320
        Shape5(1).Left = 360
        Shape5(1).Top = 5040
        Shape5(2).Left = 0
        Shape5(2).Top = 4680
        Shape5(3).Left = 720
        Shape5(3).Top = 4680
        Label1(0).Top = 4560
        Label1(0).Left = 120
        
    End If
    
    If Val(Label1(1).Caption) <= 0 Then
        Label5(0).Caption = Str(Val(Label5(0).Caption) + 1)
        Label5(3).Caption = Str(Val(Label5(3).Caption) + 1)
        Label1(1).Caption = 5
        Shape4(1).Top = 4440
        Shape4(1).Left = 12240
        Shape6(0).Left = 12480
        Shape6(0).Top = 4320
        Shape6(1).Left = 12480
        Shape6(1).Top = 5040
        Shape6(2).Left = 12120
        Shape6(2).Top = 4680
        Shape6(3).Left = 12840
        Shape6(3).Top = 4680
        Label1(1).Top = 4560
        Label1(1).Left = 12240
    End If
    
    
    '胜负判定
    If Val(Label1(6).Caption) <= 0 And Val(Label1(8).Caption) <= 0 Then
        ysl = MsgBox("右方胜利", 0, "结束")
        Timer3.Enabled = False
    End If
    
    If Val(Label1(7).Caption) <= 0 And Val(Label1(9).Caption) <= 0 Then
        zsl = MsgBox("左方胜利", 0, "结束")
        Timer3.Enabled = False
    End If

End Sub

Private Sub Timer4_Timer()
'导弹随坦克改方向bug修复
'左方
    For j = 0 To 6
        If fdd1(j) = 2 Then
            If Shape7(j).Left > 0 Then
                Shape7(j).Left = Shape7(j).Left - 150
            ElseIf Shape7(j).Left <= 0 Then
                Shape7(j).Left = 13680
                Shape7(j).Top = 2760 + 240 * j
                a(j) = False
                fdd1(j) = 4
            End If
        ElseIf fdd1(j) = 5 Then
            If Shape7(j).Top > 0 Then
                Shape7(j).Top = Shape7(j).Top - 150
            ElseIf Shape7(j).Top <= 0 Then
                Shape7(j).Left = 13680
                Shape7(j).Top = 2760 + 240 * j
                a(j) = False
                fdd1(j) = 4
            End If
        ElseIf fdd1(j) = 1 Then
            If Shape7(j).Top <= 9240 Then
                Shape7(j).Top = Shape7(j).Top + 150
            ElseIf Shape7(j).Top > 9240 Then
                Shape7(j).Left = 13680
                Shape7(j).Top = 2760 + 240 * j
                a(j) = False
                fdd1(j) = 4
            End If
        ElseIf fdd1(j) = 3 Then
            If Shape7(j).Left <= 13200 Then
                Shape7(j).Left = Shape7(j).Left + 150
            ElseIf Shape7(j).Left > 13200 Then
                Shape7(j).Left = 13680
                Shape7(j).Top = 2760 + 240 * j
                a(j) = False
                fdd1(j) = 4
            End If
        End If
    Next j
'右方
     For m = 0 To 6
        If fdd2(m) = 2 Then
            If Shape8(m).Left > 0 Then
                Shape8(m).Left = Shape8(m).Left - 150
            ElseIf Shape8(m).Left <= 0 Then
                Shape8(m).Left = 15000
                Shape8(m).Top = 2760 + 240 * m
                a2(m) = False
                fdd2(m) = 4
            End If
        ElseIf fdd2(m) = 5 Then
            If Shape8(m).Top > 0 Then
                Shape8(m).Top = Shape8(m).Top - 150
            ElseIf Shape8(m).Top <= 0 Then
                Shape8(m).Left = 15000
                Shape8(m).Top = 2760 + 240 * m
                a2(m) = False
                fdd2(m) = 4
            End If
        ElseIf fdd2(m) = 1 Then
            If Shape8(m).Top <= 9240 Then
                Shape8(m).Top = Shape8(m).Top + 150
            ElseIf Shape8(m).Top > 9240 Then
                Shape8(m).Left = 15000
                Shape8(m).Top = 2760 + 240 * m
                a2(m) = False
                fdd2(m) = 4
            End If
        ElseIf fdd2(m) = 3 Then
            If Shape8(m).Left <= 13200 Then
                Shape8(m).Left = Shape8(m).Left + 150
            ElseIf Shape8(m).Left > 13200 Then
                Shape8(m).Left = 15000
                Shape8(m).Top = 2760 + 240 * m
                a2(m) = False
                fdd2(m) = 4
            End If
        End If
    Next m
End Sub

Private Sub Timer5_Timer()
'导弹右方
If Shape6(2).Visible = True Then
    For m = 0 To 6
        If a2(m) = True Then
            If c2(m) = True Then
                Shape8(m).Left = Shape6(2).Left
                Shape8(m).Top = Shape6(2).Top
                c2(m) = False
            End If
            n = m
            If fdd2(n) = 4 Then fdd2(n) = 2
        End If
    Next m
ElseIf Shape6(0).Visible = True Then
    For m = 0 To 6
        If a2(m) = True Then
            If c2(m) = True Then
                Shape8(m).Left = Shape6(0).Left
                Shape8(m).Top = Shape6(0).Top
                c2(m) = False
            End If
            n = m
            If fdd2(n) = 4 Then fdd2(n) = 5
        End If
    Next m
ElseIf Shape6(1).Visible = True Then
    For m = 0 To 6
        If a2(m) = True Then
            If c2(m) = True Then
                Shape8(m).Left = Shape6(1).Left
                Shape8(m).Top = Shape6(1).Top
                c2(m) = False
            End If
            n = m
            If fdd2(n) = 4 Then fdd2(n) = 1
        End If
    Next m
ElseIf Shape6(3).Visible = True Then
    For m = 0 To 6
        If a2(m) = True Then
            If c2(m) = True Then
                Shape8(m).Left = Shape6(3).Left
                Shape8(m).Top = Shape6(3).Top
                c2(m) = False
            End If
            n = m
            If fdd2(n) = 4 Then fdd2(n) = 3
        End If
    Next m
End If
End Sub
