VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "会計ソフト－追加・編集"
   ClientHeight    =   5505
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8640
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********************
'FaileName :会計アプリ
'FainalDate:2020.11.24
'
'////関数////
'<- Text Box ->
'TBno : No.
'TBdate : 年月日
'*TBname : 名前
'*TBmark : 項目 **この二つはどちらか一方のみ使えるようにする。
'TBadd : 追加
'TBsub : 支出
'
'<- Button ->
'CBget：取得
'CBadd：追加
'CBedit：編集
'CBclose:閉じる
'
'<- Option ->
'OBname：名前選択
'OBmark：項目選択
'
'////その他/////
'@author nora
'@version 1.0-a
'*********************

'グローバル変数
Option Explicit
Public LastLine As Integer

'====================
'
'関数：OBname_Click
'用途：名前を入力をON/OFFする。
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
'関数：OBmark_Click
'用途：項目を入力をON/OFFする。
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
'関数：UserForm_Initialize
'用途：ユーザーフォームを開いたときに実行される。
'
'未実装：Levelの取得方法
'====================
Private Sub UserForm_Initialize()
    '初期化
    LastLine = Worksheets(sheetNameGlobal).Range("B" & Rows.Count).End(xlUp).Row
    
    'フォームに情報を追加
    Application.ScreenUpdating = False
    TBno.Value = LastLine - 3
    TBno.IMEMode = fmIMEModeAlpha
    TBdate.Value = Date
    TBdate.IMEMode = fmIMEModeAlpha
    TBname.IMEMode = fmIMEModeHiragana
    TBname.Value = Null
    'Levelの取得処理
    With LBlevel
        .AddItem "Level.1"
        .AddItem "Level.2"
        .AddItem "Level.3"
        .AddItem "Level.4"
        .AddItem "ヒーリング"
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
'関数：CBadd_Click
'用途：クリックするとsheertに内容が追加される。
'
'未実装：
'
'====================
Private Sub CBadd_Click()
On Error GoTo Exception
    '変数
    'Dim LastLine As Long    '最終行の取得
    Dim LL As Long          '最終行に１プラス
    Dim flag As Boolean        'errの検出
    Dim NoCheck As Long     'TBNoの取得
    
    '初期設定
    LastLine = Worksheets(sheetNameGlobal).Range("B" & Rows.Count).End(xlUp).Row
    LL = LastLine + 1
    
    NoCheck = TBno.Value
    flag = True
    
    'データが既に入っていないか調べる
    '全て調べる必要ある？
    Dim Err As Long         'For文
    For Err = 4 To LastLine
        If Cells(Err, 1).Value = NoCheck Then
            flag = False
            Exit For
        End If
    Next Err
    
    '表示するかと0より大きい
    If flag And NoCheck > 0 Then
        '収入/支出に値が入っている
        If (Not TBadd.Value = "" And TBsub.Value = "") Or (TBadd.Value = "" And Not TBsub.Value = "") Then
            '名前か項かを記入
            If (Not TBmark.Value = "" And TBname.Value = "") Or (TBmark.Value = "" And Not TBname.Value = "") Then
                '表の追加
                Range(Cells(LL, 1), Cells(LL, 7)).Borders.LineStyle = xlContinuous
                'データ追加
                Cells(LL, 1).Value = TBno.Value
                Cells(LL, 2).Value = TBdate.Value
                'もし、項目に文字が入っていなかったら
                If OBname.Value Then
                    Cells(LL, 3).Value = TBname.Value + " 様"
                    Cells(LL, 4).Value = LBlevel.Value
                ElseIf OBmark.Value Then
                    Cells(LL, 3).Value = TBmark.Value
                Else
                    '記載に変更がない場合。
                    MsgBox "Error:名前か、項目を記載してください。", vbOKOnly
                End If
                    Cells(LL, 5).Value = Format(TBadd.Value, "###,###,###")
                    Cells(LL, 6).Value = Format(TBsub.Value, "###,###,###")
                If Not Cells(LastLine, 7).Value = Cells(4, 7).Value Then
                    Cells(LL, 7).Value = Format((Cells(LastLine, 7).Value + Cells(LL, 5).Value) - Cells(LL, 6).Value, "###,###,###")
                Else
                    Cells(LL, 7).Value = Format(Cells(LL, 5).Value, "###,###,###")
                End If
                
                '閉じる
                Unload UserForm1
            Else
                '記載がない場合。
                MsgBox "Error:項目か、お名前を入力してください。", vbOKOnly
            End If
        Else
            '記載に変更がない場合。
            MsgBox "Error:これでは、残金の更新がありません。", vbOKOnly
        End If
    Else
        'エラー検出
        MsgBox "Error:既にデータが入っています。", vbOKOnly
    End If
    
    Exit Sub
Exception:
    MsgBox "予期しないエラーです。"
End Sub

'====================
'
'関数：CBedit_Click
'用途：クリックするとsheertの内容が編集される。
'
'未実装：
'
'====================
Private Sub CBedit_Click()
On Error GoTo Exception
    '初期化
    LastLine = Worksheets(sheetNameGlobal).Range("B" & Rows.Count).End(xlUp).Row
    
    '変数
    Dim LL As Long
    LL = LastLine + 1
    
    '取得（No.）
    Dim No As Integer
    No = TBno.Value + 4
    'TBadd.Value = No
    'TBsub.Value = LL & " " & LastLine
    
    '編集機能
    If Not (LL = No) Then
        '収入/支出に値が入っている
        If (Not TBadd.Value = "" And TBsub.Value = "") Or (TBadd.Value = "" And Not TBsub.Value = "") Then
            '名前か項かを記入
            If (Not TBmark.Value = "" And TBname.Value = "") Or (TBmark.Value = "" And Not TBname.Value = "") Then
                
                Cells(No, 2).Value = TBdate.Value
                If TBmark.Value = "" Then
                    Cells(No, 3).Value = TBname.Value + " 様"
                    Cells(No, 4).Value = LBlevel.Value
                ElseIf TBname.Value = "" Then
                    Cells(No, 3).Value = TBmark.Value
                    Cells(No, 4).Value = ""
                Else
                    '記載に変更がない場合。
                    MsgBox "Error:名前か、項目を記載してください。", vbOKOnly
                End If
                Cells(No, 5).Value = Format(TBadd.Value, "###,###,###")
                Cells(No, 6).Value = Format(TBsub.Value, "###,###,###")
                
                '計算
                Dim NoAgo As Integer
                If No > 5 Then 'セルに合わせてある（＝５）
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
                    MsgBox "error:最初は収入が必要です。", vbOKOnly
                End If
                '閉じる
                Unload UserForm1
             Else
                '記載がない場合。
                MsgBox "Error:項目か、お名前を入力してください。", vbOKOnly
            End If
        Else
            '記載に変更がない場合。
            MsgBox "Error:これでは、残金の更新がありません。", vbOKOnly
        End If
    Else
        'エラー検出
        MsgBox "Please:追加機能お使いください", vbOKOnly
    End If
    
    Exit Sub
Exception:
    MsgBox Err.Number & vbCrLf & Err.Description
End Sub

'====================
'
'関数：CBget_Click
'用途：クリックするとsheertの内容を取得します。
'
'====================
Private Sub CBget_Click()
On Error GoTo Exception
    '変数
    Dim No As Integer
    LastLine = Worksheets(sheetNameGlobal).Range("B" & Rows.Count).End(xlUp).Row
    Dim LL As Integer
    LL = LastLine + 1
    
    '初期化
    Worksheets(sheetNameGlobal).Activate
    No = TBno.Value + 4
    
    '新規の場合は取得しない
    If Not (LL = No) Then
        '番号の情報を取得
        TBdate.Value = Cells(No, 2).Value
        
        '項目か名前か
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
        'エラー検出
        MsgBox "Error:取得する値がありません。", vbOKOnly
    End If
    
    Exit Sub
Exception:
    MsgBox Err.Number & vbCrLf & Err.Description
End Sub

'====================
'
'関数：CBclose_Click
'用途：クリックするとユーザーフォームを閉じます。
'
'====================
Private Sub CBclose_Click()
On Error GoTo Exception
    Unload Me
    Exit Sub
Exception:
    MsgBox Err.Number & vbCrLf & Err.Description
End Sub
