VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "三人组"
   ClientHeight    =   6600
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8790
   LinkTopic       =   "Form1"
   ScaleHeight     =   6600
   ScaleWidth      =   8790
   StartUpPosition =   3  '窗口缺省
   Begin VB.ListBox List1 
      Height          =   3120
      Left            =   6240
      TabIndex        =   2
      Top             =   960
      Width           =   2175
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   720
      Top             =   840
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   1095
      Left            =   480
      TabIndex        =   1
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   975
      Left            =   480
      TabIndex        =   0
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "窗"
      Height          =   735
      Index           =   4
      Left            =   5880
      TabIndex        =   8
      Top             =   4440
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000011&
      Caption         =   "讲台"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   3
      Left            =   1680
      TabIndex        =   7
      Top             =   5520
      Width           =   5175
   End
   Begin VB.Label Label2 
      Caption         =   "前门"
      Height          =   495
      Left            =   2040
      TabIndex        =   6
      Top             =   4440
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      Height          =   735
      Index           =   2
      Left            =   3720
      TabIndex        =   5
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      Height          =   735
      Index           =   1
      Left            =   3720
      TabIndex        =   4
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   735
      Index           =   0
      Left            =   3960
      TabIndex        =   3
      Top             =   4200
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00008080&
      BorderWidth     =   3
      Height          =   1095
      Index           =   3
      Left            =   3480
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00008080&
      BorderWidth     =   2
      Height          =   1095
      Index           =   2
      Left            =   3480
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00008080&
      Height          =   1095
      Index           =   1
      Left            =   3480
      Top             =   4080
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Integer
Dim b As Integer

Private Sub Command1_Click()
Timer1.Enabled = False
Shape1(1).BackStyle = 0 - transparent
Shape1(2).BackStyle = 0 - transparent
Shape1(3).BackStyle = 0 - transparent
List1.Clear
End Sub

Private Sub Command2_Click()
Randomize
a = 1
If Shape1(1).BackStyle = 1 - opaque And Shape1(2).BackStyle = 1 - opaque And Shape1(3).BackStyle = 1 - opaque Then a = 0
Do While a = 1
i = Int(3 * Rnd + 1)
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
b = Int(3 * Rnd + 1)
Shape1(b).BackStyle = 0 - transparent
Else
b = i
Shape1(b).BackStyle = 0 - transparent
End If
i = Int(3 * Rnd + 1)
Shape1(i).BackStyle = 1 - opaque
b = b + 1

End Sub
