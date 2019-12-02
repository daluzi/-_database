VERSION 5.00
Begin VB.Form login 
   Caption         =   "登录"
   ClientHeight    =   6795
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9465
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   6795
   ScaleWidth      =   9465
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   3480
      TabIndex        =   3
      Top             =   2280
      Width           =   3135
   End
   Begin VB.CommandButton Command5 
      Caption         =   "取 消"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5400
      TabIndex        =   2
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton test1 
      Caption         =   "登 录"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   1
      Top             =   4440
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   3480
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   3360
      Width           =   3135
   End
   Begin VB.Label Label3 
      Caption         =   "学生登录系统"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   6
      Top             =   960
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "密 码:"
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
      Left            =   2400
      TabIndex        =   5
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "用户名:"
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
      Left            =   2520
      TabIndex        =   4
      Top             =   2280
      Width           =   855
   End
End
Attribute VB_Name = "login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command5_Click()
Unload Me
End Sub

Private Sub test1_Click()
Dim cn As New ADODB.Connection
Dim cn_str As String
cn_str = "driver=sql server;server=(local);database=da1"
cn.Open cn_str
Dim rs As New ADODB.Recordset
Dim sql As String
sql = "select * from users where u_name='" + Text1.Text + "'"
Set rs = cn.Execute(sql)
If rs.EOF Then
MsgBox "用户不存在", vbOKOnly + vbExclamation
Text1.Text = " "
Text1.SetFocus
Else
If Text2.Text = Trim(rs.Fields("u_paw")) Then
u_paw = Text1.Text
main.Show
Me.Hide
Else
MsgBox "密码错误"
Text2.Text = ""
End If
End If
End Sub

