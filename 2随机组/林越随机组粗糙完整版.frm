VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6525
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11895
   LinkTopic       =   "Form1"
   ScaleHeight     =   6525
   ScaleWidth      =   11895
   StartUpPosition =   3  '窗口缺省
   Begin VB.ListBox List1 
      Height          =   4380
      Left            =   9000
      TabIndex        =   5
      Top             =   720
      Width           =   2175
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   840
      Top             =   480
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   1215
      Left            =   360
      TabIndex        =   4
      Top             =   4080
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1095
      Left            =   360
      TabIndex        =   3
      Top             =   2640
      Width           =   1575
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
      Height          =   855
      Index           =   8
      Left            =   6240
      Top             =   1680
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808000&
      Height          =   855
      Index           =   7
      Left            =   6240
      Top             =   2880
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808000&
      Height          =   855
      Index           =   6
      Left            =   6240
      Top             =   4080
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808000&
      Height          =   855
      Index           =   5
      Left            =   4680
      Top             =   1680
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808000&
      Height          =   855
      Index           =   4
      Left            =   4680
      Top             =   2880
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808000&
      Height          =   855
      Index           =   3
      Left            =   4680
      Top             =   4080
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808000&
      Height          =   855
      Index           =   2
      Left            =   3120
      Top             =   1680
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808000&
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
End Sub

Private Sub Command2_Click()
Randomize
a = 1
e = e + 1
If Shape1(1).BackStyle = 1 - opaque And Shape1(2).BackStyle = 1 - opaque And Shape1(3).BackStyle = 1 - opaque And Shape1(4).BackStyle = 1 - opaque And Shape1(5).BackStyle = 1 - opaque And Shape1(6).BackStyle = 1 - opaque And Shape1(7).BackStyle = 1 - opaque And Shape1(8).BackStyle = 1 - opaque And Shape1(9).BackStyle = 1 - opaque Then MsgBox "请先点开始，再点一下停止就要炸了", 0, "来自林越的温馨提醒"
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
