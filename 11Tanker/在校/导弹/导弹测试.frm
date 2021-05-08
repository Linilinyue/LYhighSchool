VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7590
   ClientLeft      =   4380
   ClientTop       =   3210
   ClientWidth     =   12330
   LinkTopic       =   "Form1"
   ScaleHeight     =   7590
   ScaleWidth      =   12330
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   960
      Top             =   6120
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   3360
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   11040
      TabIndex        =   0
      Top             =   6360
      Width           =   615
   End
   Begin VB.Shape Shape1 
      Height          =   135
      Index           =   6
      Left            =   9840
      Top             =   2280
      Width           =   135
   End
   Begin VB.Shape Shape1 
      Height          =   135
      Index           =   5
      Left            =   9840
      Top             =   2040
      Width           =   135
   End
   Begin VB.Shape Shape1 
      Height          =   135
      Index           =   4
      Left            =   9840
      Top             =   1800
      Width           =   135
   End
   Begin VB.Shape Shape1 
      Height          =   135
      Index           =   3
      Left            =   9840
      Top             =   1560
      Width           =   135
   End
   Begin VB.Shape Shape1 
      Height          =   135
      Index           =   2
      Left            =   9840
      Top             =   1320
      Width           =   135
   End
   Begin VB.Shape Shape1 
      Height          =   135
      Index           =   1
      Left            =   9840
      Top             =   1080
      Width           =   135
   End
   Begin VB.Shape Shape1 
      Height          =   135
      Index           =   0
      Left            =   9840
      Top             =   840
      Width           =   135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim a(0 To 6) As Boolean
Dim b As Boolean


Private Sub command1_keydown(keycode As Integer, shift As Integer)
    If keycode = vbKeyJ Then
        b = True
        For i = 0 To 6
            If a(i) = False And b = True Then
                a(i) = True
                b = False
        End If
        Next i
    End If
    
    
End Sub

Private Sub Form_Load()
    For i = 0 To 6
        a(i) = False
    Next i
End Sub

Private Sub Timer1_Timer()
    For i = 0 To 6
        If a(i) = True Then
            If Shape1(i).Left > 0 Then
                Shape1(i).Left = Shape1(i).Left - 100
            ElseIf Shape1(i).Left <= 0 Then
                Shape1(i).Left = 9840
                a(i) = False
            End If
        End If
    Next i
End Sub
