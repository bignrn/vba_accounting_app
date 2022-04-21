VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "��v�\�t�g�|���ڂ̍폜"
   ClientHeight    =   4725
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6165
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********************
'FaileName :��v�A�v��
'FainalDate:2020.11.26
'
'////�֐�////
'<- Labele ->
'LnoIn�FNo.
'LmarkIn : ���ځE���O
'
'<- Button ->
'CBerase�F�폜
'CBclose�F��߂�
'
'////���[�U�[�֐�////
'<- Function ->
'EraseAllCell�F�폜����
'
'////���̑�/////
'@author nora
'@version 1.0
'*********************

'====================
'
'�֐��FUserForm_Initialize
'�p�r�F���[�U�[�t�H�[�����J�����Ƃ��Ɏ��s�����B
'
'�������FLevel�̎擾���@
'====================
Private Sub UserForm_Initialize()
    'MsgBox "�I�������͈͂̒��ɃZ������" & Selection.Count & "�ł�"
    '**MsgBox "�I�������͂̐擪�s��" & Selection(1).Row & "�s�ڂł�"
    'MsgBox "�I�������͂̍s����" & Selection.Rows.Count & "�s�ł�"
    '**MsgBox "�I�������͂̍ŏI�s��" & Selection(Selection.Count).Row & "�s�ڂł�"
    'MsgBox "�I�������͂̐擪���" & Selection(1).Column & "��ڂł�"
    'MsgBox "�I�������̗͂񐔂�" & Selection.Columns.Count & "��ł�"
    'MsgBox "�I�������͂̍ŏI���" & Selection(Selection.Count).Column & "��ڂł�"
    
    '�ϐ�
    Dim Fast As Integer
    Dim Last As Integer
    Dim msg As String
    Dim msgMark As String
    
    '������
    Fast = Selection(1).Row
    Last = Selection(Selection.Count).Row
    
    
    '����
    If Not (Fast < 5) Then
        If 0 < Fast Then
            If Not Fast = Last Then
                msg = "No." & Fast - 4 & "����No." & Last - 4
                msgMark = "�I��͈͂��ׂ�"
            Else
                msg = Fast - 4
                msgMark = Cells(Fast, 3).Value
            End If
            
            '�\��
            LnoIn.Caption = msg
            LmarkIn.Caption = msgMark
        Else
            '�͈͊O
            MsgBox "Error:�͈͊O�ł��B", vbOKOnly
            LnoIn.Caption = -1
            LmarkIn.Caption = "�G���[�����o����܂����B"
        End If
    Else
        MsgBox "Error:�폜�s�\�ȏꏊ�ł��B", vbOKOnly
        Call CBclose_Click
    End If
End Sub

'====================
'
'�֐��FCBclose_Click
'�p�r�F�N���b�N����ƃt�H�[������܂��B
'
'====================
Private Sub CBclose_Click()
On Error GoTo Exception
    Unload Me
    Exit Sub
Exception:
    MsgBox Err.Number & vbCrLf & Err.Description
End Sub

'====================
'
'�֐��FCBerase_Click
'�p�r�F�N���b�N�����sheert�̓��e���폜���܂��B
'
'�������F���͈̔͂��폜�A�v�Z�̂�蒼���B
'====================
Private Sub CBerase_Click()
On Error GoTo Exception
    '�ϐ�
    Dim Fast As Integer
    Dim Last As Integer
    Dim msg As String
    Dim msgMark As String
    
    '������
    Fast = Selection(1).Row
    Last = Selection(Selection.Count).Row
    
    '���f���邩
    If EraseAllCell(Fast, Last) Then
        '����
        MsgBox "�폜���܂����B", vbOKOnly
    Else
        '���s
        MsgBox "Error:�폜�Ɏ��s���܂����B", vbOKOnly
    End If
    
    '����
    Unload Me
    
    Exit Sub
Exception:
    '��O����
    MsgBox Err.Number & vbCrLf & Err.Description
End Sub

'====================
'
'�֐��FEraseAllCell()
'@param Integer FastCell
'@param Integer LastCell
'@return boolean
'�p�r�F�Z���̍폜���s���B
'
'�������F
'====================
Function EraseAllCell(FastCell As Integer, LastCell As Integer) As Boolean
On Error GoTo Exception
    '�ϐ�
    Dim flag As Boolean
    Dim LastLine As Integer
    Dim discount As Integer
    
    '������
    flag = False
    LastLine = Worksheets(sheetNameGlobal).Range("B" & Rows.Count).End(xlUp).Row
    
    '��O����
    If LmarkIn.Caption = "�G���[�����o����܂����B" Then
        EraseAllCell = flag
        Exit Function
    ElseIf FastCell < 5 Then
        EraseAllCell = flag
        Exit Function
    End If
    
    '���C������
    If LastCell <= LastLine Then
        '�폜����
        Application.ScreenUpdating = False
        
        For discount = LastCell To FastCell Step -1
            'MsgBox discount & "�f���[�g"
            Cells(discount, 7).Delete
            Cells(discount, 6).Delete
            Cells(discount, 5).Delete
            Cells(discount, 4).Delete
            Cells(discount, 3).Delete
            Cells(discount, 2).Delete
        Next discount
        
        Application.ScreenUpdating = True
        
        '�Ď擾
        LastLineNew = Worksheets(sheetNameGlobal).Range("B" & Rows.Count).End(xlUp).Row
        
        '�v�Z
        For discount = FastCell To LastLineNew - 1
            If Not discount - 1 = 4 Then
                Cells(discount, 7).Value = Format((Cells(discount - 1, 7).Value + Cells(discount, 5).Value) - Cells(discount, 6).Value, "###,###,###")
            Else
                Cells(discount, 7).Value = Format(Cells(discount, 5).Value, "###,###,###")
            End If
            'MsgBox discount & "�v�Z"
        Next discount
        
        'No������
        For discount = LastLine To LastLineNew + 1 Step -1
            Cells(discount, 1).Delete
        Next discount
        
        flag = True
    Else
        '���s
        flag = False
    End If
    
    '�߂�l
    EraseAllCell = flag

    Exit Function
Exception:
    '��O����
    MsgBox Err.Number & vbCrLf & Err.Description
End Function
