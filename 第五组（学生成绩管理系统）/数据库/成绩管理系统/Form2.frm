VERSION 5.00
Begin VB.Form main 
   Caption         =   "���ѧ����Ϣ"
   ClientHeight    =   4875
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   7335
   LinkTopic       =   "Form2"
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   4875
   ScaleWidth      =   7335
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Menu xt 
      Caption         =   "ϵͳ"
   End
   Begin VB.Menu xx_cx 
      Caption         =   "ѧ����Ϣ��ѯ"
      Begin VB.Menu cx_xx 
         Caption         =   "���ѧ��"
         Begin VB.Menu rg 
            Caption         =   "�������ѧ�뼼��1441"
            Begin VB.Menu xs_xx 
               Caption         =   "ѧ����Ϣ"
            End
            Begin VB.Menu lr_cj 
               Caption         =   "¼��ɼ�"
            End
            Begin VB.Menu xs_cj 
               Caption         =   "ѧ���ɼ�"
            End
         End
      End
      Begin VB.Menu tj_xx 
         Caption         =   "���ѧ��"
      End
      Begin VB.Menu xg_xx 
         Caption         =   "�޸�����"
      End
      Begin VB.Menu sc_xx 
         Caption         =   "ɾ��ѧ��"
         Begin VB.Menu am_sc 
            Caption         =   "������ɾ��"
         End
         Begin VB.Menu ax_sc 
            Caption         =   "��ѧ��ɾ��"
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
scxx.Label2.Caption = "����"
End Sub

Private Sub ax_sc_Click()
scxx.Show
scxx.Label2.Caption = "ѧ��"
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
