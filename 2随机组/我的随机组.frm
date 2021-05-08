VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "林越随机组粗糙完整版"
   ClientHeight    =   6525
   ClientLeft      =   7665
   ClientTop       =   4155
   ClientWidth     =   11895
   LinkTopic       =   "Form1"
   ScaleHeight     =   6525
   ScaleWidth      =   11895
   Begin VB.CommandButton Command6 
      Caption         =   "三人组重选"
      Height          =   735
      Left            =   8160
      TabIndex        =   10
      Top             =   5400
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   375
      Left            =   6240
      TabIndex        =   9
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   495
      Left            =   4680
      TabIndex        =   8
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   3120
      TabIndex        =   7
      Top             =   1080
      Width           =   975
   End
   Begin VB.ListBox List1 
      Height          =   1680
      Left            =   7920
      TabIndex        =   5
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   840
      Top             =   480
   End
   Begin VB.CommandButton Command2 
      Caption         =   "停止"
      Height          =   1215
      Left            =   360
      TabIndex        =   4
      Top             =   4080
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "开始"
      Height          =   1095
      Left            =   360
      TabIndex        =   3
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "9"
      Height          =   495
      Index           =   8
      Left            =   6360
      TabIndex        =   19
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      Height          =   495
      Index           =   7
      Left            =   6360
      TabIndex        =   18
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      Height          =   495
      Index           =   6
      Left            =   6360
      TabIndex        =   17
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      Height          =   495
      Index           =   5
      Left            =   4800
      TabIndex        =   16
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      Height          =   495
      Index           =   4
      Left            =   4800
      TabIndex        =   15
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      Height          =   495
      Index           =   3
      Left            =   4920
      TabIndex        =   14
      Top             =   4320
      Width           =   615
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      Height          =   495
      Index           =   2
      Left            =   3240
      TabIndex        =   13
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      Height          =   495
      Index           =   1
      Left            =   3240
      TabIndex        =   12
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   495
      Index           =   0
      Left            =   3360
      TabIndex        =   11
      Top             =   4320
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "选择三人组在哪一列："
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808000&
      Height          =   855
      Index           =   9
      Left            =   3120
      Top             =   4080
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808000&
      BorderWidth     =   9
      Height          =   855
      Index           =   8
      Left            =   6240
      Top             =   1680
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808000&
      BorderWidth     =   8
      Height          =   855
      Index           =   7
      Left            =   6240
      Top             =   2880
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808000&
      BorderWidth     =   7
      Height          =   855
      Index           =   6
      Left            =   6240
      Top             =   4080
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808000&
      BorderWidth     =   6
      Height          =   855
      Index           =   5
      Left            =   4680
      Top             =   1680
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808000&
      BorderWidth     =   5
      Height          =   855
      Index           =   4
      Left            =   4680
      Top             =   2880
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808000&
      BorderWidth     =   4
      Height          =   855
      Index           =   3
      Left            =   4680
      Top             =   4080
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808000&
      BorderWidth     =   3
      Height          =   855
      Index           =   2
      Left            =   3120
      Top             =   1680
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808000&
      BorderWidth     =   2
      Height          =   855
      Index           =   1
      Left            =   3120
      Top             =   2880
      Width           =   975
   End
   Begin VB.Shape Shape11 
      Height          =   1095
      Left            =   360
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "窗"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8040
      TabIndex        =   2
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "前门"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   1
      Top             =   5640
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "讲台"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   0
      Top             =   5760
      Width           =   1455
   End
   Begin VB.Shape Shape10 
      Height          =   735
      Left            =   2280
      Top             =   5640
      Width           =   5295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim a As Integer
Dim b As Integer
Dim c As Integer
Dim d As Integer
Dim e As Integer
Dim f As String
Dim g As Integer
Dim h As String
Dim s As Integer

Private Sub Command1_Click()
Timer1.Enabled = False
Shape1(1).BackStyle = 0 - transparent
Shape1(2).BackStyle = 0 - transparent
Shape1(3).BackStyle = 0 - transparent
Shape1(4).BackStyle = 0 - transparent
Shape1(5).BackStyle = 0 - transparent
Shape1(6).BackStyle = 0 - transparent
Shape1(7).BackStyle = 0 - transparent
Shape1(8).BackStyle = 0 - transparent
Shape1(9).BackStyle = 0 - transparent
List1.Clear
e = 0
Command4.Enabled = True
Command3.Enabled = True
Command5.Enabled = True

End Sub

Private Sub Command2_Click()
Randomize
a = 1
If Shape1(1).BackStyle = 1 - opaque And Shape1(2).BackStyle = 1 - opaque And Shape1(3).BackStyle = 1 - opaque And Shape1(4).BackStyle = 1 - opaque And Shape1(5).BackStyle = 1 - opaque And Shape1(6).BackStyle = 1 - opaque And Shape1(7).BackStyle = 1 - opaque And Shape1(8).BackStyle = 1 - opaque And Shape1(9).BackStyle = 1 - opaque Then a = 0
Do While a = 1
i = Int(9 * Rnd + 1)
If Shape1(i).BackStyle = 1 - opaque Then
a = 1
Else
a = 0
End If
Loop
If a = 0 Then Shape1(i).BackStyle = 1 - opaque
List1.AddItem Shape1(i).BorderWidth

End Sub

Private Sub Command3_Click()
Shape1(9).BackStyle = 1 - opaque
Shape1(1).BackStyle = 1 - opaque
Shape1(2).BackStyle = 1 - opaque
Command4.Enabled = False
Command5.Enabled = False

End Sub

Private Sub Command4_Click()
Shape1(5).BackStyle = 1 - opaque
Shape1(4).BackStyle = 1 - opaque
Shape1(3).BackStyle = 1 - opaque
Command5.Enabled = False
Command3.Enabled = False

End Sub

Private Sub Command5_Click()
Shape1(6).BackStyle = 1 - opaque
Shape1(8).BackStyle = 1 - opaque
Shape1(7).BackStyle = 1 - opaque
Command4.Enabled = False
Command3.Enabled = False
End Sub

Private Sub Command6_Click()
f = "三人组.exe"
g = Shell(f, vbNormalFocus)
End Sub

Private Sub Form_Load()
f = "三人组.exe"
g = Shell(f, vbNormalFocus)
h = "登记本.exe"
s = Shell(h, vbNormalFocus)

End Sub

Private Sub Timer1_Timer()

b = 0
If b = 0 Then
b = Int(9 * Rnd + 1)
Shape1(b).BackStyle = 0 - transparent
Else
b = i
Shape1(b).BackStyle = 0 - transparent
End If
i = Int(9 * Rnd + 1)
Shape1(i).BackStyle = 1 - opaque
b = b + 1

End Sub
