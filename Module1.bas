Attribute VB_Name = "Module1"
'*********************
'Main Class
'FaileName :Module1
'FainalDate:2020.11.25
'
'@ather nora
'@version 1.0
'
'�Q�l����
'https://www.sejuku.net/blog/35484
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
On Error GoTo Exception
    UserForm1.Show
    Exit Sub
Exception:
        MsgBox Err.Number & vbCrLf & Err.Description
End Sub

'====================
'
'�֐��FEraseUserFoemOPEN_Click
'�p�r�F�N���b�N����ƃ��[�U�[�t�H�[���������オ��
'
'====================
Sub EraseUserFormOPEN_Click()
On Error GoTo Exception
    UserForm2.Show
    Exit Sub
Exception:
        MsgBox Err.Number & vbCrLf & Err.Description
End Sub

'====================
'
'�֐��FBackSheet_Click
'�p�r�F�N���b�N������O�̃V�[�g�ɖ߂�
'
'====================
Sub BackSheet_Click()
On Error GoTo Exception
    Worksheets(sheetNameGlobal).Activate
    Exit Sub
Exception:
        MsgBox Err.Number & vbCrLf & Err.Description
End Sub

