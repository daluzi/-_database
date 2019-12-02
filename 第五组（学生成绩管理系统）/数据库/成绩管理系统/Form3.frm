VERSION 5.00
Begin VB.Form llxs 
   Caption         =   "学生个人信息查询"
   ClientHeight    =   6810
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8925
   LinkTopic       =   "Form3"
   Picture         =   "Form3.frx":0000
   ScaleHeight     =   6810
   ScaleWidth      =   8925
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command4 
      Caption         =   "上一条"
      Height          =   495
      Left            =   3120
      TabIndex        =   12
      Top             =   5160
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "下一条"
      Height          =   495
      Left            =   4560
      TabIndex        =   11
      Top             =   5160
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "最后一条"
      Height          =   495
      Left            =   6120
      TabIndex        =   10
      Top             =   5160
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "首记录"
      Height          =   495
      Left            =   1680
      TabIndex        =   9
      Top             =   5160
      Width           =   975
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   5280
      TabIndex        =   8
      Top             =   3240
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   1920
      TabIndex        =   6
      Top             =   3240
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   5280
      TabIndex        =   4
      Top             =   1800
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1920
      TabIndex        =   2
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   8880
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Label Label5 
      Caption         =   "专业："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4320
      TabIndex        =   7
      Top             =   3360
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "性别："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   5
      Top             =   3360
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "姓名："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4320
      TabIndex        =   3
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "学号："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "浏览学生个人信息"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Top             =   480
      Width           =   2655
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   9240
      Y1              =   1320
      Y2              =   1320
   End
End
Attribute VB_Name = "llxs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rst As New ADODB.Recordset

Private Sub Command1_Click()
    rst.MoveFirst
display
End Sub

Private Sub Command2_Click()
rst.MoveLast
    display
End Sub

Private Sub Command3_Click()
rst.MoveNext
If rst.EOF Then
    MsgBox "已经是最后一条记录了", vbOKOnly + vbExclamation
rst.MoveLast
End If
    display
End Sub

Private Sub Command4_Click()
rst.MovePrevious
If rst.BOF Then
MsgBox "已经是第一条记录了", vbOKOnly + vbExclamation
rst.MoveFirst
End If
display
End Sub

Private Sub Form_Load()
Dim str As String
str = "select * from student"
Set rst = chaxun(str)
display
End Sub
Private Sub display()
Text1.Text = rst.Fields(0)
Text2.Text = rst.Fields(1)
Text3.Text = rst.Fields(2)
Text4.Text = rst.Fields(3)
End Sub

