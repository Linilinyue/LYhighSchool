VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "林越，测试跑道"
   ClientHeight    =   7440
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12450
   LinkTopic       =   "Form1"
   ScaleHeight     =   7440
   ScaleWidth      =   12450
   StartUpPosition =   3  '窗口缺省
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   36
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3660
      Index           =   1
      Left            =   6240
      TabIndex        =   9
      Top             =   1200
      Width           =   855
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   36
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3660
      Index           =   3
      Left            =   3120
      TabIndex        =   6
      Top             =   1200
      Width           =   855
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   36
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3660
      Index           =   2
      Left            =   3960
      TabIndex        =   5
      Top             =   1200
      Width           =   855
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   36
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3660
      Index           =   1
      Left            =   4800
      TabIndex        =   4
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "开始"
      Height          =   975
      Left            =   480
      TabIndex        =   3
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "停止"
      Height          =   855
      Left            =   480
      TabIndex        =   2
      Top             =   5040
      Width           =   1215
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   36
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3660
      Index           =   3
      Left            =   7920
      TabIndex        =   1
      Top             =   1200
      Width           =   855
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   36
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3660
      Index           =   2
      Left            =   7080
      TabIndex        =   0
      Top             =   1200
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   5880
      Top             =   360
   End
   Begin VB.Label Label3 
      Caption         =   " "
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   36
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   3
      Left            =   7920
      TabIndex        =   15
      Top             =   5040
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   " "
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   36
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   3
      Left            =   3120
      TabIndex        =   14
      Top             =   5040
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   " "
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   36
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   4080
      TabIndex        =   13
      Top             =   5040
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   36
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   4920
      TabIndex        =   12
      Top             =   5040
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   " "
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   36
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   2
      Left            =   7200
      TabIndex        =   11
      Top             =   5040
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   36
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   1
      Left            =   6240
      TabIndex        =   10
      Top             =   5040
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   720
      TabIndex        =   8
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "分数："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   7
      Top             =   1320
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Integer
Dim i As Integer
Dim ch As Integer
Dim ji As Integer
Dim fen As Integer
Dim a1(1 To 3) As Integer
Dim a2(1 To 3) As Integer
Dim b1(1 To 3) As Integer
Dim b2(1 To 3) As Integer
Dim c1(1 To 3) As Integer
Dim c2(1 To 3) As Integer
Dim d1(1 To 3) As Integer
Dim d2(1 To 3) As Integer
Dim e1(1 To 3) As Integer
Dim e2(1 To 3) As Integer
Dim f1(1 To 3) As Integer
Dim f2(1 To 3) As Integer

Private Sub command1_KeyDown(KeyCode As Integer, shift As Integer)
 If KeyCode = vbKeyJ Then
  If ch > 1 Then
   ch = ch - 1
   Label3(ch).Caption = 1
   Label4(ch).Caption = 1
   Label3(ch + 1).Caption = " "
   Label4(ch + 1).Caption = " "
  End If
 ElseIf KeyCode = vbKeyK Then
  If ch < 3 Then
   ch = ch + 1
   Label3(ch).Caption = 1
   Label4(ch).Caption = 1
   Label3(ch - 1).Caption = " "
   Label4(ch - 1).Caption = " "
  End If
 End If
End Sub

Private Sub Command2_Click()
Timer1.Enabled = False
End Sub

Private Sub Command1_Click()
Timer1.Enabled = True
Label2.Caption = 0
End Sub

Private Sub Form_Load()
Randomize
ch = 1
End Sub

Private Sub Timer1_Timer()

'跑道部分:
ji = ji + 1
If ji Mod 3 = 0 Then
 i = Int(Rnd * 3 + 1)
 a1(i) = 1
 a2(i) = 1
 For i = 1 To 3
   If a1(i) <> 1 Then
     a1(i) = Int(Rnd * 2 + 1)
   End If
 Next i
 
  For i = 1 To 3
   If a2(i) <> 1 Then
     a2(i) = Int(Rnd * 2 + 1)
   End If
 Next i


 For a = 1 To 3
   If a1(a) = 1 Then
     List1(a).AddItem " ", 0
   Else
      List1(a).AddItem "#", 0
    End If
   If a2(a) = 1 Then
     List2(a).AddItem " ", 0
   Else
      List2(a).AddItem "#", 0
    End If
 Next a


Else
 For a = 1 To 3
  List1(a).AddItem " ", 0
  List2(a).AddItem " ", 0
 Next a
End If

'存储障碍物的位置:
For i = 1 To 3
  f1(i) = e1(i)
  e1(i) = d1(i)
  d1(i) = c1(i)
  c1(i) = b1(i)
  b1(i) = a1(i)
  
  f2(i) = e2(i)
  e2(i) = d2(i)
  d2(i) = c2(i)
  c2(i) = b2(i)
  b2(i) = a2(i)
Next i

'记分数:
For i = 1 To 3
  If a1(i) = 2 Then
    fen = fen + 1
    If Timer1.Interval >= 200 Then
    Timer1.Interval = Timer1.Interval - 5
    End If
  End If
  If a2(i) = 2 Then
  fen = fen + 1
    If Timer1.Interval >= 200 Then
      Timer1.Interval = Timer1.Interval - 5
    End If
  End If
Next i
Label2.Caption = fen

 For a = 1 To 3
   a1(a) = 0
   a2(a) = 0
 Next a

'死亡部分
For i = 1 To 3
  If f1(i) = 2 And Label3(i).Caption = "1" Then
  b = MsgBox("结束了", 0, "提醒")
  Timer1.Enabled = False
  Timer1.Interval = 500
  End If
   If f2(i) = 2 And Label4(i).Caption = "1" Then
  b = MsgBox("结束了", 0, "提醒")
  Timer1.Enabled = False
  Timer1.Interval = 500
  End If
Next i
End Sub
