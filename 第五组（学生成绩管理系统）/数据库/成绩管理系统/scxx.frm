VERSION 5.00
Begin VB.Form scxx 
   Caption         =   "ɾ����Ϣ"
   ClientHeight    =   4665
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10305
   LinkTopic       =   "Form1"
   Picture         =   "scxx.frx":0000
   ScaleHeight     =   4665
   ScaleWidth      =   10305
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command2 
      Caption         =   "ȡ��"
      Height          =   375
      Left            =   5400
      TabIndex        =   4
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ȷ��"
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   3000
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   1
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "ɾ��ѧ����Ϣ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   0
      Top             =   720
      Width           =   2055
   End
End
Attribute VB_Name = "scxx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim rs As ADODB.Recordset
    Dim str As String
    If Label2.Caption = "����" Then
    str = "select * from student where  stu_name='" + Text1.Text + "'"
    Else
    str = "select * from student where  stu_id='" + Text1.Text + "'"
    End If
    Set rs = chaxun(str)
    If rs.EOF Then
     MsgBox "û��Ҫɾ������Ϣ", vbOKOnly + vbExclamation
    Else
        While Not rs.EOF
          rs.Delete
          rs.MoveNext
    Wend
    MsgBox "ɾ���ɹ�", vbOKOnly + vbExclamation
    End If
End Sub

Private Sub Command2_Click()
Me.Hide
End Sub
