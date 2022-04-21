Attribute VB_Name = "vertion20"
'*********************
'Vertion Updata
'FaileName :vertion20
'FainalDate:2021.2.22
'
'@ather nora
'*********************
'定数
Public Const vertion = "ver.2.0"

'*****************
'用途:vertionのアップデート記載
'*****************
Public Function NewVertionUpData()
    If Not Cells(1, 8).Value = vertion Then
        Cells(1, 8).Value = "ver.2.0"
        Cells(5, 9).Value = ""
        Worksheets(sheetNameGlobal2).Activate 'シート2をアクティブ
        Cells(1, 5).Value = "ver.1.0-a"
        Cells(3, 4).Value = "更新時刻"
        Cells(3, 4).Borders(xlEdgeRight).LineStyle = xlContinuous '更新時刻の右に罫線追加
        MsgBox "アップデータが完了しました。", vbOKOnly
    End If
End Function
