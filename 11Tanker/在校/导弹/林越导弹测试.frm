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
      Left            =   480
      Top             =   6480
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   7560
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "按J发射"
      Height          =   495
      Left            =   11040
      TabIndex        =   0
      Top             =   6360
      Width           =   855
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
Dim a(0 To 6) As Boolean        '运动，判断是否到底
Dim b As Boolean                '每次发一个
Dim c(0 To 6) As Boolean        '从tanker位置发射



Private Sub command1_keydown(keycode As Integer, shift As Integer)
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
    
    
End Sub

Private Sub Form_Load()
    For i = 0 To 6
        a(i) = False
    Next i
End Sub

Private Sub Timer1_Timer()
    Command2.Top = Command2.Top + 12
    If Command2.Top >= 7200 Then Command2.Top = 0
    For i = 0 To 6
        If a(i) = True Then
            If c(i) = True Then
                Shape1(i).Left = Command2.Left
                Shape1(i).Top = Command2.Top
                c(i) = False
            End If
            
            If Shape1(i).Left > 0 Then
                Shape1(i).Left = Shape1(i).Left - 150
            ElseIf Shape1(i).Left <= 0 Then
                Shape1(i).Left = 9840
                Shape1(i).Top = 840 + 240 * i
                a(i) = False
            End If
        End If
    Next i
End Sub
