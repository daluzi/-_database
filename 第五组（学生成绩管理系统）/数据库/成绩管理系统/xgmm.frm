VERSION 5.00
Begin VB.Form xgmm 
   Caption         =   "�޸�����"
   ClientHeight    =   5595
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10080
   LinkTopic       =   "Form1"
   Picture         =   "xgmm.frx":0000
   ScaleHeight     =   5595
   ScaleWidth      =   10080
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command2 
      Caption         =   "ȷ��"
      Height          =   375
      Left            =   3960
      TabIndex        =   8
      Top             =   4320
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   3960
      TabIndex        =   7
      Top             =   2280
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   3960
      TabIndex        =   5
      Top             =   3120
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ȷ��"
      Height          =   375
      Left            =   7200
      TabIndex        =   3
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label Label4 
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
      Left            =   2640
      TabIndex        =   6
      Top             =   2280
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "ȷ������"
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
      Left            =   2640
      TabIndex        =   4
      Top             =   3120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "ԭ����"
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
      Left            =   2640
      TabIndex        =   1
      Top             =   1440
      Width           =   975
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   10080
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label1 
      Caption         =   "�޸ĸ�������"
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
      Top             =   360
      Width           =   2055
   End
End
Attribute VB_Name = "xgmm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Private Sub Command1_Click()
Dim str As String
    str = "select * from users where u_paw='" + u_paw + "'"
    Set rs = chaxun(str)
    If Trim(rs.Fields(1)) = Text1.Text Then
    Label2.Visible = False
    Label3.Visible = True
    Label4.Visible = True
    Text1.Visible = False
    Text2.Visible = True
    Text3.Visible = True
    Command1.Visible = False
    Command2.Visible = True
    Else
     MsgBox "�����������!", vbOKOnly + vbExclamation
    End If
End Sub

Private Sub Command2_Click()
If Text2.Text = Text3.Text Then
    rs.Fields(1) = Text2.Text
    rs.Update
      MsgBox "�޸ĳɹ���", vbOKOnly + vbExclamation
    Else
      MsgBox "������������벻��ͬ��", vbOKOnly + vbExclamation
    End If
End Sub

