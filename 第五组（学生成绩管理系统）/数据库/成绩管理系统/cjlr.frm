VERSION 5.00
Begin VB.Form cjlr 
   Caption         =   "Form1"
   ClientHeight    =   5730
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7740
   LinkTopic       =   "Form1"
   Picture         =   "cjlr.frx":0000
   ScaleHeight     =   5730
   ScaleWidth      =   7740
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command1 
      Caption         =   "ȷ��¼��"
      Height          =   375
      Left            =   3600
      TabIndex        =   13
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton Command8 
      Caption         =   "���һ��"
      Height          =   375
      Left            =   5280
      TabIndex        =   8
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton Command7 
      Caption         =   "��һ��"
      Height          =   375
      Left            =   2640
      TabIndex        =   7
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton Command6 
      Caption         =   "��һ��"
      Height          =   375
      Left            =   3960
      TabIndex        =   6
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "����"
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   3960
      Width           =   855
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   4800
      TabIndex        =   4
      Top             =   2040
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   2040
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "�ɼ���"
      Height          =   375
      Left            =   3720
      TabIndex        =   12
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "�Ա�"
      Height          =   375
      Left            =   840
      TabIndex        =   11
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "������"
      Height          =   375
      Left            =   3720
      TabIndex        =   10
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "ѧ�ţ�"
      Height          =   375
      Left            =   720
      TabIndex        =   9
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "ѧ���ɼ�¼��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "cjlr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rst As New ADODB.Recordset

Private Sub Command1_Click()
Dim rst As New ADODB.Recordset
Dim str As String
str = "select*from grade"
Set rst = chaxun(str)
rst.AddNew
rst.Fields(3) = Text4.Text
rst.Update
MsgBox "��ӳɹ�", vbOKOnly + vbExclamation
End Sub

Private Sub Command5_Click()
    rst.MoveFirst
display
End Sub

Private Sub Command6_Click()
rst.MoveNext
If rst.EOF Then
    MsgBox "�Ѿ������һ����¼��", vbOKOnly + vbExclamation
rst.MoveLast
End If
    display
End Sub

Private Sub Command7_Click()
rst.MovePrevious
If rst.BOF Then
MsgBox "�Ѿ��ǵ�һ����¼��", vbOKOnly + vbExclamation
rst.MoveFirst
End If
display
End Sub

Private Sub Command8_Click()
rst.MoveLast
    display
End Sub

Private Sub Form_Load()
Dim str As String
str = "select * from grade"
Set rst = chaxun(str)
display
End Sub
Private Sub display()
Text1.Text = rst.Fields(0)
Text2.Text = rst.Fields(1)
Text3.Text = rst.Fields(2)
End Sub

