VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "友谊的小船？"
   ClientHeight    =   4530
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10365
   LinkTopic       =   "Form1"
   ScaleHeight     =   4530
   ScaleWidth      =   10365
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "更改"
      Height          =   975
      Left            =   3600
      TabIndex        =   1
      Top             =   3120
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   1215
      Left            =   4680
      TabIndex        =   0
      Top             =   1200
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As String
Dim b As Integer

Private Sub Command1_Click()
a = InputBox("你最好的朋友是谁？", "问答")
Label1.Caption = a
If a = "林越" Then MsgBox "这还差不多", 0, "嗯。。"
If a <> "林越" Then
 b = MsgBox("不玩了！", 0, "再见！")
 If b = vbOK Then End
End If
End Sub

Private Sub Form_Load()
a = InputBox("你最好的朋友是谁？", "问答")
Label1.Caption = a
If a <> "林越" Then MsgBox "不跟你好了！赶紧点更改", 0, "哼"
If a = "林越" Then MsgBox "嘿嘿你好啊", 0, "嗯。。"
End Sub
