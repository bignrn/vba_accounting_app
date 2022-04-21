Attribute VB_Name = "AdditionalImplemantation"
'*********************
'Additional implementation
'FaileName :AdditionalInplemantation
'FainalDate:2021.2.22
'
'@ather nora
'@version 1.0-a
'*********************
'�萔
Const sheetName1 = sheetNameGlobal
Const sheetName2 = sheetNameGlobal2
Const inLine = 6 '�擾��̏����l
Const outLine = 7 '�擾��̏����l
Const sumValue = 0 '�c���v�Z
Dim ����(12) As Long
Dim �x�o(12) As Long
Dim flg As Boolean
'====================
'
'�֐��FCalculation_Click()
'�p�r�F�N���b�N����ƌ����Ƃɂ܂Ƃ߂�
'
'====================
Sub Calculation_Click()
On Error GoTo Exception
    '�ϐ�
    Dim m As Integer
    Dim ���t As Long
    Dim �N As Long
    Dim �����N As Long
    Dim interval As Integer
    Dim �ŏI�s1 As Long
    Dim �ŏI�s2 As Long
    �ŏI�s1 = Worksheets(sheetName1).Range("B" & Rows.Count).End(xlUp).Row
    �ŏI�s2 = Worksheets(sheetName2).Range("B" & Rows.Count).End(xlUp).Row
    interval = 0
    flg = False
    
    'version�A�b�v�f�[�g
    Call NewVertionUpData
    
    '���C��
    Application.ScreenUpdating = False '�`���~
    Worksheets(sheetName1).Activate '�V�[�g�P���A�N�e�B�u
    
    '������
    �����N = year(Cells(5, 2).Value)
    
    '������ŏI�s�܂Ōv�Z����
    For m = 5 To �ŏI�s1
        ���t = Cells(m, 2).Value  '�V�[�g�P�̓��t����擾
        �N = year(���t)
        
        If �����N = �N Then
            '�v�Z���\�b�h
            Call addCellsValue(���t, m)
        Else
            '�\�����\�b�h
            Call drowSheet(interval, �����N)
            '������
            For i = 1 To 12
                ����(i) = 0
                �x�o(i) = 0
            Next i
            '�ēx������
            Application.ScreenUpdating = False '�`���~
            Worksheets(sheetName1).Activate '�V�[�g�P���A�N�e�B�u
            '����
            �����N = �����N + 1
            interval = interval + 12
            
            '�v�Z���\�b�h
            Call addCellsValue(���t, m)
        End If
    Next m
    '�o�̓��\�b�h
    Call drowSheet(interval, �����N)
    
    '�z��̏�����
    For i = 1 To 12
        ����(i) = 0
        �x�o(i) = 0
    Next i
    
    '�X�V���ԕ\��
    Cells(3, 5).Value = Date & "/" & Time
    
    Cells(�ŏI�s + 1, 1).Select

    Exit Sub
Exception:
    '��O����
    MsgBox Err.Number & vbCrLf & Err.Description
End Sub
'*****************
'�p�r:�v�Z���ʂ��o��
'*****************
Private Function drowSheet(interval As Integer, yearName As Long)
    Application.ScreenUpdating = True '�`�拖��
    Worksheets(sheetName2).Activate '�V�[�g�Q�̃A�N�e�B�u
    
    Dim standardLine As Long 'for�̏����l
    Dim standardLineLast As Long 'for�̏����l
    standardLine = 5
    standardLineLast = 16
    
    Cells(standardLine + interval, 1).Value = yearName '�N�o��
    
    '�v�Z���ʂ��o��
    Dim index As Integer
    Dim indexIntarval As Integer
    For i = (standardLine + interval) To (standardLineLast + interval)
        Range(Cells(i, 1), Cells(i, 5)).Borders.LineStyle = xlContinuous
        index = (i - interval)
        indexIntarval = (standardLine - 1)
        Cells(i, 2).Value = i - (indexIntarval + interval) '��
        Cells(i, 3).Value = ����(index - indexIntarval)
        Cells(i, 4).Value = �x�o(index - indexIntarval)
        Cells(i, 5).Value = balanceCalc(Cells(i, 3).Value, Cells(i, 4).Value, i)
    Next i
End Function
'*****************
'�p�r:�v�Z���\�b�h(������)
'*****************
Private Function addCellsValue(dateData As Long, m As Integer)
    Select Case Month(dateData) '�����ƂɌv�Z���ʂ�ϐ��֑��
        Case 1
            ����(1) = ����(1) + Cells(m, inLine).Value
            �x�o(1) = �x�o(1) + Cells(m, outLine).Value
        Case 2
            ����(2) = ����(2) + Cells(m, inLine).Value
            �x�o(2) = �x�o(2) + Cells(m, outLine).Value
        Case 3
            ����(3) = ����(3) + Cells(m, inLine).Value
            �x�o(3) = �x�o(3) + Cells(m, outLine).Value
        Case 4
            ����(4) = ����(4) + Cells(m, inLine).Value
            �x�o(4) = �x�o(4) + Cells(m, outLine).Value
        Case 5
            ����(5) = ����(5) + Cells(m, inLine).Value
            �x�o(5) = �x�o(5) + Cells(m, outLine).Value
        Case 6
            ����(6) = ����(6) + Cells(m, inLine).Value
            �x�o(6) = �x�o(6) + Cells(m, outLine).Value
        Case 7
            ����(7) = ����(7) + Cells(m, inLine).Value
            �x�o(7) = �x�o(7) + Cells(m, outLine).Value
        Case 8
            ����(8) = ����(8) + Cells(m, inLine).Value
            �x�o(8) = �x�o(8) + Cells(m, outLine).Value
        Case 9
            ����(9) = ����(9) + Cells(m, inLine).Value
            �x�o(9) = �x�o(9) + Cells(m, outLine).Value
        Case 10
            ����(10) = ����(10) + Cells(m, inLine).Value
            �x�o(10) = �x�o(10) + Cells(m, outLine).Value
        Case 11
            ����(11) = ����(11) + Cells(m, inLine).Value
            �x�o(11) = �x�o(11) + Cells(m, outLine).Value
        Case 12
            ����(12) = ����(12) + Cells(m, inLine).Value
            �x�o(12) = �x�o(12) + Cells(m, outLine).Value
    End Select
End Function
'*****************
'�p�r:�v�Z���\�b�h2(�c��)
'*****************
Private Function balanceCalc(add As Long, subb As Long, ByVal i As Long) As Long
    '�ϐ�
    Dim calc As Long
    
    '�c���v�Z
    If flg Then
        calc = (add + Cells(i - 1, 5).Value) - subb
    Else
        '���ڂ͎c�����Ȃ�����v�Z����
        calc = add - subb
        flg = True
    End If
    
    If calc < 0 Then
        MsgBox "�}�C�i�X�̒l���o�܂����B�m�F���Ă�������", vbOKOnly
    End If
    
    balanceCalc = calc
End Function
