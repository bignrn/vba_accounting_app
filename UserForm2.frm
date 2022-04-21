VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "会計ソフト－項目の削除"
   ClientHeight    =   4725
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6165
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********************
'FaileName :会計アプリ
'FainalDate:2020.11.26
'
'////関数////
'<- Labele ->
'LnoIn：No.
'LmarkIn : 項目・名前
'
'<- Button ->
'CBerase：削除
'CBclose：やめる
'
'////ユーザー関数////
'<- Function ->
'EraseAllCell：削除する
'
'////その他/////
'@author nora
'@version 1.0
'*********************

'====================
'
'関数：UserForm_Initialize
'用途：ユーザーフォームを開いたときに実行される。
'
'未実装：Levelの取得方法
'====================
Private Sub UserForm_Initialize()
    'MsgBox "選択した範囲の中にセル数は" & Selection.Count & "個です"
    '**MsgBox "選択した範の先頭行は" & Selection(1).Row & "行目です"
    'MsgBox "選択した範の行数は" & Selection.Rows.Count & "行です"
    '**MsgBox "選択した範の最終行は" & Selection(Selection.Count).Row & "行目です"
    'MsgBox "選択した範の先頭列は" & Selection(1).Column & "列目です"
    'MsgBox "選択した範の列数は" & Selection.Columns.Count & "列です"
    'MsgBox "選択した範の最終列は" & Selection(Selection.Count).Column & "列目です"
    
    '変数
    Dim Fast As Integer
    Dim Last As Integer
    Dim msg As String
    Dim msgMark As String
    
    '初期化
    Fast = Selection(1).Row
    Last = Selection(Selection.Count).Row
    
    
    '処理
    If Not (Fast < 5) Then
        If 0 < Fast Then
            If Not Fast = Last Then
                msg = "No." & Fast - 4 & "からNo." & Last - 4
                msgMark = "選択範囲すべて"
            Else
                msg = Fast - 4
                msgMark = Cells(Fast, 3).Value
            End If
            
            '表示
            LnoIn.Caption = msg
            LmarkIn.Caption = msgMark
        Else
            '範囲外
            MsgBox "Error:範囲外です。", vbOKOnly
            LnoIn.Caption = -1
            LmarkIn.Caption = "エラーが検出されました。"
        End If
    Else
        MsgBox "Error:削除不可能な場所です。", vbOKOnly
        Call CBclose_Click
    End If
End Sub

'====================
'
'関数：CBclose_Click
'用途：クリックするとフォームを閉じます。
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
'関数：CBerase_Click
'用途：クリックするとsheertの内容を削除します。
'
'未実装：その範囲を削除、計算のやり直し。
'====================
Private Sub CBerase_Click()
On Error GoTo Exception
    '変数
    Dim Fast As Integer
    Dim Last As Integer
    Dim msg As String
    Dim msgMark As String
    
    '初期化
    Fast = Selection(1).Row
    Last = Selection(Selection.Count).Row
    
    '反映するか
    If EraseAllCell(Fast, Last) Then
        '成功
        MsgBox "削除しました。", vbOKOnly
    Else
        '失敗
        MsgBox "Error:削除に失敗しました。", vbOKOnly
    End If
    
    '閉じる
    Unload Me
    
    Exit Sub
Exception:
    '例外処理
    MsgBox Err.Number & vbCrLf & Err.Description
End Sub

'====================
'
'関数：EraseAllCell()
'@param Integer FastCell
'@param Integer LastCell
'@return boolean
'用途：セルの削除を行う。
'
'未実装：
'====================
Function EraseAllCell(FastCell As Integer, LastCell As Integer) As Boolean
On Error GoTo Exception
    '変数
    Dim flag As Boolean
    Dim LastLine As Integer
    Dim discount As Integer
    
    '初期化
    flag = False
    LastLine = Worksheets(sheetNameGlobal).Range("B" & Rows.Count).End(xlUp).Row
    
    '例外処理
    If LmarkIn.Caption = "エラーが検出されました。" Then
        EraseAllCell = flag
        Exit Function
    ElseIf FastCell < 5 Then
        EraseAllCell = flag
        Exit Function
    End If
    
    'メイン処理
    If LastCell <= LastLine Then
        '削除処理
        Application.ScreenUpdating = False
        
        For discount = LastCell To FastCell Step -1
            'MsgBox discount & "デリート"
            Cells(discount, 7).Delete
            Cells(discount, 6).Delete
            Cells(discount, 5).Delete
            Cells(discount, 4).Delete
            Cells(discount, 3).Delete
            Cells(discount, 2).Delete
        Next discount
        
        Application.ScreenUpdating = True
        
        '再取得
        LastLineNew = Worksheets(sheetNameGlobal).Range("B" & Rows.Count).End(xlUp).Row
        
        '計算
        For discount = FastCell To LastLineNew - 1
            If Not discount - 1 = 4 Then
                Cells(discount, 7).Value = Format((Cells(discount - 1, 7).Value + Cells(discount, 5).Value) - Cells(discount, 6).Value, "###,###,###")
            Else
                Cells(discount, 7).Value = Format(Cells(discount, 5).Value, "###,###,###")
            End If
            'MsgBox discount & "計算"
        Next discount
        
        'Noを消す
        For discount = LastLine To LastLineNew + 1 Step -1
            Cells(discount, 1).Delete
        Next discount
        
        flag = True
    Else
        '失敗
        flag = False
    End If
    
    '戻り値
    EraseAllCell = flag

    Exit Function
Exception:
    '例外処理
    MsgBox Err.Number & vbCrLf & Err.Description
End Function
