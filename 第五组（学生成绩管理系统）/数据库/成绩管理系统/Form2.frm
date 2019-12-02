VERSION 5.00
Begin VB.Form main 
   Caption         =   "浏览学生信息"
   ClientHeight    =   4875
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   7335
   LinkTopic       =   "Form2"
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   4875
   ScaleWidth      =   7335
   StartUpPosition =   3  '窗口缺省
   Begin VB.Menu xt 
      Caption         =   "系统"
   End
   Begin VB.Menu xx_cx 
      Caption         =   "学生信息查询"
      Begin VB.Menu cx_xx 
         Caption         =   "浏览学生"
         Begin VB.Menu rg 
            Caption         =   "计算机科学与技术1441"
            Begin VB.Menu xs_xx 
               Caption         =   "学生信息"
            End
            Begin VB.Menu lr_cj 
               Caption         =   "录入成绩"
            End
            Begin VB.Menu xs_cj 
               Caption         =   "学生成绩"
            End
         End
      End
      Begin VB.Menu tj_xx 
         Caption         =   "添加学生"
      End
      Begin VB.Menu xg_xx 
         Caption         =   "修改密码"
      End
      Begin VB.Menu sc_xx 
         Caption         =   "删除学生"
         Begin VB.Menu am_sc 
            Caption         =   "按姓名删除"
         End
         Begin VB.Menu ax_sc 
            Caption         =   "按学号删除"
         End
      End
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
Form3.Show
End Sub

Private Sub am_sc_Click()
scxx.Show
scxx.Label2.Caption = "姓名"
End Sub

Private Sub ax_sc_Click()
scxx.Show
scxx.Label2.Caption = "学号"
End Sub

Private Sub lr_cj_Click()
cjlr.Show
End Sub

Private Sub tj_xx_Click()
tjxx.Show
End Sub

Private Sub xg_xx_Click()
xgmm.Show
End Sub

Private Sub xs_cj_Click()
cjcx.Show
End Sub

Private Sub xs_xx_Click()
llxs.Show
End Sub
