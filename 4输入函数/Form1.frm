VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "�����С����"
   ClientHeight    =   4530
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10365
   LinkTopic       =   "Form1"
   ScaleHeight     =   4530
   ScaleWidth      =   10365
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command1 
      Caption         =   "����"
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
a = InputBox("����õ�������˭��", "�ʴ�")
Label1.Caption = a
If a = "��Խ" Then MsgBox "�⻹���", 0, "�š���"
If a <> "��Խ" Then
 b = MsgBox("�����ˣ�", 0, "�ټ���")
 If b = vbOK Then End
End If
End Sub

Private Sub Form_Load()
a = InputBox("����õ�������˭��", "�ʴ�")
Label1.Caption = a
If a <> "��Խ" Then MsgBox "��������ˣ��Ͻ������", 0, "��"
If a = "��Խ" Then MsgBox "�ٺ���ð�", 0, "�š���"
End Sub
