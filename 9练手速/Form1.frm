VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "练手速202小越子     分数：0/50"
   ClientHeight    =   5715
   ClientLeft      =   4860
   ClientTop       =   3285
   ClientWidth     =   7905
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   7905
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFF80&
      Height          =   375
      Left            =   3720
      MaskColor       =   &H000080FF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2640
      Width           =   375
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   800
      Left            =   7560
      Top             =   0
   End
   Begin VB.Label Label1 
      Caption         =   "50"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Integer
Dim b As Integer


Private Sub Command2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    a = a + 1
    Form1.Caption = "练手速202小越子     分数：" + Str(a) + "/50"
End Sub

Private Sub Form_Load()
Command2.Enabled = False
End Sub



Private Sub Label1_DblClick()
    Timer1.Enabled = True
    Command2.Enabled = True
    b = 0
    Form1.Caption = "练手速202小越子     分数：0/50"
End Sub

Private Sub Timer1_Timer()
    Randomize
    Command2.Top = Int(Rnd * (Form1.Height - Command2.Height * 2))

    Randomize
    Command2.Left = Int(Rnd * (Form1.Width - Command2.Width * 2))
    
    b = b + 1
    Label1.Caption = 50 - b
    If b = 50 Then
        Command2.Enabled = False
        Timer1.Enabled = False
        a = 0
    End If
End Sub
