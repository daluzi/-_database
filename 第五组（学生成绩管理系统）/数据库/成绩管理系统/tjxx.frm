VERSION 5.00
Begin VB.Form tjxx 
   Caption         =   "添加学生信息"
   ClientHeight    =   5520
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10560
   LinkTopic       =   "Form1"
   Palette         =   "tjxx.frx":0000
   Picture         =   "tjxx.frx":30C0D
   ScaleHeight     =   5520
   ScaleWidth      =   10560
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "取消添加"
      Height          =   375
      Left            =   6120
      TabIndex        =   10
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定添加"
      Height          =   375
      Left            =   3480
      TabIndex        =   9
      Top             =   4440
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2880
      TabIndex        =   3
      Top             =   1920
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   6240
      TabIndex        =   2
      Top             =   1920
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   2880
      TabIndex        =   1
      Top             =   3360
      Width           =   1815
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   6240
      TabIndex        =   0
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Line Line1 
      X1              =   960
      X2              =   10200
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label Label1 
      Caption         =   "添加学生个人信息"
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
      Left            =   4200
      TabIndex        =   8
      Top             =   600
      Width           =   2655
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
      Left            =   1920
      TabIndex        =   7
      Top             =   2040
      Width           =   855
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
      Left            =   5280
      TabIndex        =   6
      Top             =   2040
      Width           =   855
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
      Left            =   1920
      TabIndex        =   5
      Top             =   3480
      Width           =   855
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
      Left            =   5280
      TabIndex        =   4
      Top             =   3480
      Width           =   855
   End
End
Attribute VB_Name = "tjxx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim rst As New ADODB.Recordset
Dim str As String
str = "select*from student"
Set rst = chaxun(str)
rst.AddNew
rst.Fields(0) = Text1.Text
rst.Fields(1) = Text2.Text
rst.Fields(2) = Text3.Text
rst.Fields(3) = Text4.Text
rst.Update
MsgBox "添加成功", vbOKOnly + vbExclamation
End Sub

Private Sub Command2_Click()
Me.Hide
End Sub

