Attribute VB_Name = "vertion20"
'*********************
'Vertion Updata
'FaileName :vertion20
'FainalDate:2021.2.22
'
'@ather nora
'*********************
'�萔
Public Const vertion = "ver.2.0"

'*****************
'�p�r:vertion�̃A�b�v�f�[�g�L��
'*****************
Public Function NewVertionUpData()
    If Not Cells(1, 8).Value = vertion Then
        Cells(1, 8).Value = "ver.2.0"
        Cells(5, 9).Value = ""
        Worksheets(sheetNameGlobal2).Activate '�V�[�g2���A�N�e�B�u
        Cells(1, 5).Value = "ver.1.0-a"
        Cells(3, 4).Value = "�X�V����"
        Cells(3, 4).Borders(xlEdgeRight).LineStyle = xlContinuous '�X�V�����̉E�Ɍr���ǉ�
        MsgBox "�A�b�v�f�[�^���������܂����B", vbOKOnly
    End If
End Function
