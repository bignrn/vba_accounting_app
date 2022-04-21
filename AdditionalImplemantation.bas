Attribute VB_Name = "AdditionalImplemantation"
'*********************
'Additional implementation
'FaileName :AdditionalInplemantation
'FainalDate:2021.2.22
'
'@ather nora
'@version 1.0-a
'*********************
'定数
Const sheetName1 = sheetNameGlobal
Const sheetName2 = sheetNameGlobal2
Const inLine = 6 '取得列の初期値
Const outLine = 7 '取得列の初期値
Const sumValue = 0 '残高計算
Dim 収入(12) As Long
Dim 支出(12) As Long
Dim flg As Boolean
'====================
'
'関数：Calculation_Click()
'用途：クリックすると月ごとにまとめる
'
'====================
Sub Calculation_Click()
On Error GoTo Exception
    '変数
    Dim m As Integer
    Dim 日付 As Long
    Dim 年 As Long
    Dim 初期年 As Long
    Dim interval As Integer
    Dim 最終行1 As Long
    Dim 最終行2 As Long
    最終行1 = Worksheets(sheetName1).Range("B" & Rows.Count).End(xlUp).Row
    最終行2 = Worksheets(sheetName2).Range("B" & Rows.Count).End(xlUp).Row
    interval = 0
    flg = False
    
    'versionアップデート
    Call NewVertionUpData
    
    'メイン
    Application.ScreenUpdating = False '描画停止
    Worksheets(sheetName1).Activate 'シート１をアクティブ
    
    '初期化
    初期年 = year(Cells(5, 2).Value)
    
    '頭から最終行まで計算する
    For m = 5 To 最終行1
        日付 = Cells(m, 2).Value  'シート１の日付から取得
        年 = year(日付)
        
        If 初期年 = 年 Then
            '計算メソッド
            Call addCellsValue(日付, m)
        Else
            '表示メソッド
            Call drowSheet(interval, 初期年)
            '初期化
            For i = 1 To 12
                収入(i) = 0
                支出(i) = 0
            Next i
            '再度初期化
            Application.ScreenUpdating = False '描画停止
            Worksheets(sheetName1).Activate 'シート１をアクティブ
            '次へ
            初期年 = 初期年 + 1
            interval = interval + 12
            
            '計算メソッド
            Call addCellsValue(日付, m)
        End If
    Next m
    '出力メソッド
    Call drowSheet(interval, 初期年)
    
    '配列の初期化
    For i = 1 To 12
        収入(i) = 0
        支出(i) = 0
    Next i
    
    '更新時間表示
    Cells(3, 5).Value = Date & "/" & Time
    
    Cells(最終行 + 1, 1).Select

    Exit Sub
Exception:
    '例外処理
    MsgBox Err.Number & vbCrLf & Err.Description
End Sub
'*****************
'用途:計算結果を出力
'*****************
Private Function drowSheet(interval As Integer, yearName As Long)
    Application.ScreenUpdating = True '描画許可
    Worksheets(sheetName2).Activate 'シート２のアクティブ
    
    Dim standardLine As Long 'forの初期値
    Dim standardLineLast As Long 'forの初期値
    standardLine = 5
    standardLineLast = 16
    
    Cells(standardLine + interval, 1).Value = yearName '年出力
    
    '計算結果を出力
    Dim index As Integer
    Dim indexIntarval As Integer
    For i = (standardLine + interval) To (standardLineLast + interval)
        Range(Cells(i, 1), Cells(i, 5)).Borders.LineStyle = xlContinuous
        index = (i - interval)
        indexIntarval = (standardLine - 1)
        Cells(i, 2).Value = i - (indexIntarval + interval) '月
        Cells(i, 3).Value = 収入(index - indexIntarval)
        Cells(i, 4).Value = 支出(index - indexIntarval)
        Cells(i, 5).Value = balanceCalc(Cells(i, 3).Value, Cells(i, 4).Value, i)
    Next i
End Function
'*****************
'用途:計算メソッド(月ごと)
'*****************
Private Function addCellsValue(dateData As Long, m As Integer)
    Select Case Month(dateData) '月ごとに計算結果を変数へ代入
        Case 1
            収入(1) = 収入(1) + Cells(m, inLine).Value
            支出(1) = 支出(1) + Cells(m, outLine).Value
        Case 2
            収入(2) = 収入(2) + Cells(m, inLine).Value
            支出(2) = 支出(2) + Cells(m, outLine).Value
        Case 3
            収入(3) = 収入(3) + Cells(m, inLine).Value
            支出(3) = 支出(3) + Cells(m, outLine).Value
        Case 4
            収入(4) = 収入(4) + Cells(m, inLine).Value
            支出(4) = 支出(4) + Cells(m, outLine).Value
        Case 5
            収入(5) = 収入(5) + Cells(m, inLine).Value
            支出(5) = 支出(5) + Cells(m, outLine).Value
        Case 6
            収入(6) = 収入(6) + Cells(m, inLine).Value
            支出(6) = 支出(6) + Cells(m, outLine).Value
        Case 7
            収入(7) = 収入(7) + Cells(m, inLine).Value
            支出(7) = 支出(7) + Cells(m, outLine).Value
        Case 8
            収入(8) = 収入(8) + Cells(m, inLine).Value
            支出(8) = 支出(8) + Cells(m, outLine).Value
        Case 9
            収入(9) = 収入(9) + Cells(m, inLine).Value
            支出(9) = 支出(9) + Cells(m, outLine).Value
        Case 10
            収入(10) = 収入(10) + Cells(m, inLine).Value
            支出(10) = 支出(10) + Cells(m, outLine).Value
        Case 11
            収入(11) = 収入(11) + Cells(m, inLine).Value
            支出(11) = 支出(11) + Cells(m, outLine).Value
        Case 12
            収入(12) = 収入(12) + Cells(m, inLine).Value
            支出(12) = 支出(12) + Cells(m, outLine).Value
    End Select
End Function
'*****************
'用途:計算メソッド2(残高)
'*****************
Private Function balanceCalc(add As Long, subb As Long, ByVal i As Long) As Long
    '変数
    Dim calc As Long
    
    '残高計算
    If flg Then
        calc = (add + Cells(i - 1, 5).Value) - subb
    Else
        '一回目は残高がないから計算だけ
        calc = add - subb
        flg = True
    End If
    
    If calc < 0 Then
        MsgBox "マイナスの値が出ました。確認してください", vbOKOnly
    End If
    
    balanceCalc = calc
End Function
