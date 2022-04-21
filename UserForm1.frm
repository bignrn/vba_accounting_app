VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "��v�\�t�g�|�ǉ��E�ҏW"
   ClientHeight    =   5505
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8640
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********************
'FaileName :��v�A�v��
'FainalDate:2020.11.24
'
'////�֐�////
'<- Text Box ->
'TBno : No.
'TBdate : �N����
'*TBname : ���O
'*TBmark : ���� **���̓�͂ǂ��炩����̂ݎg����悤�ɂ���B
'TBadd : �ǉ�
'TBsub : �x�o
'
'<- Button ->
'CBget�F�擾
'CBadd�F�ǉ�
'CBedit�F�ҏW
'CBclose:����
'
'<- Option ->
'OBname�F���O�I��
'OBmark�F���ڑI��
'
'////���̑�/////
'@author nora
'@version 1.0-a
'*********************

'�O���[�o���ϐ�
Option Explicit
Public LastLine As Integer

'====================
'
'�֐��FOBname_Click
'�p�r�F���O����͂�ON/OFF����B
'
'====================
Private Sub OBname_Click()
    TBname.Enabled = True
    LBlevel.Enabled = True
    TBmark.Enabled = False
    TBname.BackColor = RGB(255, 255, 255)
    LBlevel.BackColor = RGB(255, 255, 255)
    TBmark.BackColor = RGB(128, 128, 128)
    OBmark.Value = False
    OBname.Value = True
End Sub

'====================
'
'�֐��FOBmark_Click
'�p�r�F���ڂ���͂�ON/OFF����B
'
'====================
Private Sub OBmark_Click()
    TBname.Enabled = False
    LBlevel.Enabled = False
    TBmark.Enabled = True
    TBname.BackColor = RGB(128, 128, 128)
    LBlevel.BackColor = RGB(128, 128, 128)
    TBmark.BackColor = RGB(255, 255, 255)
    OBmark.Value = True
    OBname.Value = False
End Sub

'====================
'
'�֐��FUserForm_Initialize
'�p�r�F���[�U�[�t�H�[�����J�����Ƃ��Ɏ��s�����B
'
'�������FLevel�̎擾���@
'====================
Private Sub UserForm_Initialize()
    '������
    LastLine = Worksheets(sheetNameGlobal).Range("B" & Rows.Count).End(xlUp).Row
    
    '�t�H�[���ɏ���ǉ�
    Application.ScreenUpdating = False
    TBno.Value = LastLine - 3
    TBno.IMEMode = fmIMEModeAlpha
    TBdate.Value = Date
    TBdate.IMEMode = fmIMEModeAlpha
    TBname.IMEMode = fmIMEModeHiragana
    TBname.Value = Null
    'Level�̎擾����
    With LBlevel
        .AddItem "Level.1"
        .AddItem "Level.2"
        .AddItem "Level.3"
        .AddItem "Level.4"
        .AddItem "�q�[�����O"
    End With
    LBlevel.IMEMode = fmIMEModeAlpha
    TBmark.IMEMode = fmIMEModeHiragana
    TBmark.Value = Null
    TBadd.Value = Null
    TBadd.IMEMode = fmIMEModeAlpha
    TBsub.Value = Null
    TBsub.IMEMode = fmIMEModeAlpha
    
    Application.ScreenUpdating = True
    Worksheets(sheetNameGlobal).Activate
    'Cells(LastLine + 1, 2).Select
End Sub

'====================
'
'�֐��FCBadd_Click
'�p�r�F�N���b�N�����sheert�ɓ��e���ǉ������B
'
'�������F
'
'====================
Private Sub CBadd_Click()
On Error GoTo Exception
    '�ϐ�
    'Dim LastLine As Long    '�ŏI�s�̎擾
    Dim LL As Long          '�ŏI�s�ɂP�v���X
    Dim flag As Boolean        'err�̌��o
    Dim NoCheck As Long     'TBNo�̎擾
    
    '�����ݒ�
    LastLine = Worksheets(sheetNameGlobal).Range("B" & Rows.Count).End(xlUp).Row
    LL = LastLine + 1
    
    NoCheck = TBno.Value
    flag = True
    
    '�f�[�^�����ɓ����Ă��Ȃ������ׂ�
    '�S�Ē��ׂ�K�v����H
    Dim Err As Long         'For��
    For Err = 4 To LastLine
        If Cells(Err, 1).Value = NoCheck Then
            flag = False
            Exit For
        End If
    Next Err
    
    '�\�����邩��0���傫��
    If flag And NoCheck > 0 Then
        '����/�x�o�ɒl�������Ă���
        If (Not TBadd.Value = "" And TBsub.Value = "") Or (TBadd.Value = "" And Not TBsub.Value = "") Then
            '���O���������L��
            If (Not TBmark.Value = "" And TBname.Value = "") Or (TBmark.Value = "" And Not TBname.Value = "") Then
                '�\�̒ǉ�
                Range(Cells(LL, 1), Cells(LL, 7)).Borders.LineStyle = xlContinuous
                '�f�[�^�ǉ�
                Cells(LL, 1).Value = TBno.Value
                Cells(LL, 2).Value = TBdate.Value
                '�����A���ڂɕ����������Ă��Ȃ�������
                If OBname.Value Then
                    Cells(LL, 3).Value = TBname.Value + " �l"
                    Cells(LL, 4).Value = LBlevel.Value
                ElseIf OBmark.Value Then
                    Cells(LL, 3).Value = TBmark.Value
                Else
                    '�L�ڂɕύX���Ȃ��ꍇ�B
                    MsgBox "Error:���O���A���ڂ��L�ڂ��Ă��������B", vbOKOnly
                End If
                    Cells(LL, 5).Value = Format(TBadd.Value, "###,###,###")
                    Cells(LL, 6).Value = Format(TBsub.Value, "###,###,###")
                If Not Cells(LastLine, 7).Value = Cells(4, 7).Value Then
                    Cells(LL, 7).Value = Format((Cells(LastLine, 7).Value + Cells(LL, 5).Value) - Cells(LL, 6).Value, "###,###,###")
                Else
                    Cells(LL, 7).Value = Format(Cells(LL, 5).Value, "###,###,###")
                End If
                
                '����
                Unload UserForm1
            Else
                '�L�ڂ��Ȃ��ꍇ�B
                MsgBox "Error:���ڂ��A�����O����͂��Ă��������B", vbOKOnly
            End If
        Else
            '�L�ڂɕύX���Ȃ��ꍇ�B
            MsgBox "Error:����ł́A�c���̍X�V������܂���B", vbOKOnly
        End If
    Else
        '�G���[���o
        MsgBox "Error:���Ƀf�[�^�������Ă��܂��B", vbOKOnly
    End If
    
    Exit Sub
Exception:
    MsgBox "�\�����Ȃ��G���[�ł��B"
End Sub

'====================
'
'�֐��FCBedit_Click
'�p�r�F�N���b�N�����sheert�̓��e���ҏW�����B
'
'�������F
'
'====================
Private Sub CBedit_Click()
On Error GoTo Exception
    '������
    LastLine = Worksheets(sheetNameGlobal).Range("B" & Rows.Count).End(xlUp).Row
    
    '�ϐ�
    Dim LL As Long
    LL = LastLine + 1
    
    '�擾�iNo.�j
    Dim No As Integer
    No = TBno.Value + 4
    'TBadd.Value = No
    'TBsub.Value = LL & " " & LastLine
    
    '�ҏW�@�\
    If Not (LL = No) Then
        '����/�x�o�ɒl�������Ă���
        If (Not TBadd.Value = "" And TBsub.Value = "") Or (TBadd.Value = "" And Not TBsub.Value = "") Then
            '���O���������L��
            If (Not TBmark.Value = "" And TBname.Value = "") Or (TBmark.Value = "" And Not TBname.Value = "") Then
                
                Cells(No, 2).Value = TBdate.Value
                If TBmark.Value = "" Then
                    Cells(No, 3).Value = TBname.Value + " �l"
                    Cells(No, 4).Value = LBlevel.Value
                ElseIf TBname.Value = "" Then
                    Cells(No, 3).Value = TBmark.Value
                    Cells(No, 4).Value = ""
                Else
                    '�L�ڂɕύX���Ȃ��ꍇ�B
                    MsgBox "Error:���O���A���ڂ��L�ڂ��Ă��������B", vbOKOnly
                End If
                Cells(No, 5).Value = Format(TBadd.Value, "###,###,###")
                Cells(No, 6).Value = Format(TBsub.Value, "###,###,###")
                
                '�v�Z
                Dim NoAgo As Integer
                If No > 5 Then '�Z���ɍ��킹�Ă���i���T�j
                    For NoAgo = No - 1 To LastLine - 1
                        Cells(No, 7).Value = Format((Cells(NoAgo, 7).Value + Cells(No, 5).Value) - Cells(No, 6).Value, "###,###,###")
                        No = No + 1
                    Next
                ElseIf (Not TBadd.Value = "" And TBsub.Value = "") Then
                    Cells(No, 7).Value = Cells(No, 5).Value
                    For NoAgo = No To LastLine - 1
                        Cells(No + 1, 7).Value = Format((Cells(NoAgo, 7).Value + Cells(No + 1, 5).Value) - Cells(No + 1, 6).Value, "###,###,###")
                        No = No + 1
                    Next
                Else
                    MsgBox "error:�ŏ��͎������K�v�ł��B", vbOKOnly
                End If
                '����
                Unload UserForm1
             Else
                '�L�ڂ��Ȃ��ꍇ�B
                MsgBox "Error:���ڂ��A�����O����͂��Ă��������B", vbOKOnly
            End If
        Else
            '�L�ڂɕύX���Ȃ��ꍇ�B
            MsgBox "Error:����ł́A�c���̍X�V������܂���B", vbOKOnly
        End If
    Else
        '�G���[���o
        MsgBox "Please:�ǉ��@�\���g����������", vbOKOnly
    End If
    
    Exit Sub
Exception:
    MsgBox Err.Number & vbCrLf & Err.Description
End Sub

'====================
'
'�֐��FCBget_Click
'�p�r�F�N���b�N�����sheert�̓��e���擾���܂��B
'
'====================
Private Sub CBget_Click()
On Error GoTo Exception
    '�ϐ�
    Dim No As Integer
    LastLine = Worksheets(sheetNameGlobal).Range("B" & Rows.Count).End(xlUp).Row
    Dim LL As Integer
    LL = LastLine + 1
    
    '������
    Worksheets(sheetNameGlobal).Activate
    No = TBno.Value + 4
    
    '�V�K�̏ꍇ�͎擾���Ȃ�
    If Not (LL = No) Then
        '�ԍ��̏����擾
        TBdate.Value = Cells(No, 2).Value
        
        '���ڂ����O��
        If Not Cells(No, 4) = "" Then
            Call OBname_Click
            TBmark.Value = ""
            TBname.Value = Cells(No, 3).Value
            LBlevel.Value = Cells(No, 4).Value
        Else
            Call OBmark_Click
            TBname.Value = ""
            LBlevel.Value = ""
            TBmark.Value = Cells(No, 3).Value
        End If
        TBadd.Value = Cells(No, 5).Value
        TBsub.Value = Cells(No, 6).Value
    Else
        '�G���[���o
        MsgBox "Error:�擾����l������܂���B", vbOKOnly
    End If
    
    Exit Sub
Exception:
    MsgBox Err.Number & vbCrLf & Err.Description
End Sub

'====================
'
'�֐��FCBclose_Click
'�p�r�F�N���b�N����ƃ��[�U�[�t�H�[������܂��B
'
'====================
Private Sub CBclose_Click()
On Error GoTo Exception
    Unload Me
    Exit Sub
Exception:
    MsgBox Err.Number & vbCrLf & Err.Description
End Sub
