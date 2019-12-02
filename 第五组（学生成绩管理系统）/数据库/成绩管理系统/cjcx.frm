VERSION 5.00
Begin VB.Form cjcx 
   Caption         =   "学生成绩查询"
   ClientHeight    =   5145
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11070
   LinkTopic       =   "Form1"
   Picture         =   "cjcx.frx":0000
   ScaleHeight     =   5145
   ScaleWidth      =   11070
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command8 
      Caption         =   "最后一条"
      Height          =   375
      Left            =   6960
      TabIndex        =   8
      Top             =   3360
      Width           =   855
   End
   Begin VB.CommandButton Command7 
      Caption         =   "上一条"
      Height          =   375
      Left            =   4320
      TabIndex        =   7
      Top             =   3360
      Width           =   855
   End
   Begin VB.CommandButton Command6 
      Caption         =   "下一条"
      Height          =   375
      Left            =   5640
      TabIndex        =   6
      Top             =   3360
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "首条"
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   3360
      Width           =   855
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   6720
      TabIndex        =   4
      Top             =   2280
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Top             =   2280
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   6720
      TabIndex        =   2
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "成绩："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5760
      TabIndex        =   12
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "性别："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      TabIndex        =   11
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "姓名："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5760
      TabIndex        =   10
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "学号："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      TabIndex        =   9
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "学生成绩查询"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4680
      TabIndex        =   0
      Top             =   600
      Width           =   1455
   End
End
Attribute VB_Name = "cjcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rst As New ADODB.Recordset

Private Sub Command5_Click()
 rst.MoveFirst
display
End Sub

Private Sub Command6_Click()
rst.MoveNext
If rst.EOF Then
    MsgBox "已经是最后一条记录了", vbOKOnly + vbExclamation
rst.MoveLast
End If
    display
End Sub

Private Sub Command7_Click()
rst.MovePrevious
If rst.BOF Then
MsgBox "已经是第一条记录了", vbOKOnly + vbExclamation
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
Text4.Text = rst.Fields(3)
End Sub

