Attribute VB_Name = "Module1"
'*********************
'Main Class
'FaileName :Module1
'FainalDate:2020.11.25
'
'@ather nora
'@version 1.0
'*********************

'�O���[�o���ϐ�
Option Explicit
Public Const sheetNameGlobal = "Sheet1"
Public Const sheetNameGlobal2 = "Sheet2"
Public vertion As String
Public LastLine As Integer

'====================
'
'�֐��FUserFoemOPEN_Click
'�p�r�F�N���b�N����ƃ��[�U�[�t�H�[���������オ��
'
'====================
Sub UserFormOPEN_Click()
    UserForm1.Show
End Sub

'====================
'
'�֐��FEraseUserFoemOPEN_Click
'�p�r�F�N���b�N����ƃ��[�U�[�t�H�[���������オ��
'
'====================
Sub EraseUserFormOPEN_Click()
    UserForm2.Show
End Sub

'====================
'
'�֐��FBackSheet_Click
'�p�r�F�N���b�N������O�̃V�[�g�ɖ߂�
'
'====================
Sub BackSheet_Click()
    Worksheets(sheetNameGlobal).Activate
End Sub

