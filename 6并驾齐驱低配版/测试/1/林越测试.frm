VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "林越，测试跑道"
   ClientHeight    =   6585
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10155
   LinkTopic       =   "Form1"
   ScaleHeight     =   6585
   ScaleWidth      =   10155
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   975
      Left            =   480
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   855
      Left            =   2880
      TabIndex        =   3
      Top             =   5400
      Width           =   1815
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4020
      Index           =   3
      Left            =   6600
      TabIndex        =   2
      Top             =   840
      Width           =   855
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4020
      Index           =   2
      Left            =   5760
      TabIndex        =   1
      Top             =   840
      Width           =   855
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4020
      Index           =   1
      Left            =   4920
      TabIndex        =   0
      Top             =   840
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   6000
      Top             =   360
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Integer
Dim i As Integer
Dim ji As Integer
Dim a1(1 To 3) As Integer


Private Sub Command1_Click()
Timer1.Enabled = False
End Sub

Private Sub Command2_Click()
Timer1.Enabled = True
End Sub

Private Sub Form_Load()
Randomize
End Sub

Private Sub Timer1_Timer()
ji = ji + 1
If ji Mod 3 = 0 Then
 i = Int(Rnd * 3 + 1)
 a1(i) = 1
 For i = 1 To 3
   If a1(i) <> 1 Then
     a1(i) = Int(Rnd * 2 + 1)
   End If
 Next i

 For a = 1 To 3
   If a1(a) = 1 Then
     List1(a).AddItem " ", 0
   Else
      List1(a).AddItem "#", 0
    End If
 Next a

 For a = 1 To 3
   a1(a) = 0
 Next a
Else
 For a = 1 To 3
  List1(a).AddItem " ", 0
 Next a
End If
End Sub
